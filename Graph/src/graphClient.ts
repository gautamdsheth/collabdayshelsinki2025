// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import { AzureOpenAI } from 'openai';
import dotenv from 'dotenv';
import { ChatCompletionMessageParam } from 'openai/resources';

dotenv.config();

/**
 * This class is a wrapper for the Microsoft Graph API.
 * See: https://developer.microsoft.com/en-us/graph for more information.
 */
export class GraphClient {
    graphClient: Client;
    _token: string;

    constructor(token: string) {
        if (!token || !token.trim()) {
            throw new Error('GraphClient: Invalid token received.');
        }

        this._token = token;

        // Get an authenticated Microsoft Graph client using the token issued to the user.
        this.graphClient = Client.init({
            authProvider: (done) => {
                done(null, this._token); // First parameter takes an error if you can't get an access token.
            }
        });
    }

    public async getUserBySkills(query: string): Promise<{ displayName: string; workEmail?: string }[]> {
        // Extract skills (if possible), search SharePoint for each skill in parallel,
        // then combine and dedupe results.
        const skills = await this.extractSkills(query);
        const uniqueSkills = Array.from(new Set(skills.map(s => s.trim()).filter(Boolean)));

        // Run searches in parallel and return the combined, deduped results
        const resultsArrays = await Promise.all(uniqueSkills.map(s => this.runSearchFor(s)));
        return this.combineResults(resultsArrays);
    }

    // Extract skill strings from the user's query using Azure OpenAI when configured.
    private async extractSkills(query: string): Promise<string[]> {
        const openaiEndpoint = process.env.AZURE_OPENAI_ENDPOINT;
        const openaiApiKey = process.env.AZURE_OPENAI_API_KEY;
        const openaiDeployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;

        if (!openaiEndpoint || !openaiApiKey) {
            console.warn('AZURE_OPENAI_ENDPOINT or AZURE_OPENAI_API_KEY not set; skipping extraction and using provided query');
            return [query];
        }

        try {
            const apiVersion = process.env.OPENAI_API_VERSION || '2024-10-21';
            const client = new AzureOpenAI({ endpoint: openaiEndpoint, apiKey: openaiApiKey, deployment: openaiDeployment, apiVersion });

            const messages: ChatCompletionMessageParam[] = [
                { role: 'system', content: 'You are a strict extractor. When asked for skills, return ONLY a JSON array of skill strings, with no extra text. Example: ["Strategic Thinking","Team Building"]. If none found, return []' },
                { role: 'user', content: `Extract all the skills mentioned in the following prompt. Return ONLY a JSON array of strings or an empty array. Prompt:\n\n${query}` }
            ];

            const completion = await client.chat.completions.create({
                model: openaiDeployment!,
                messages: messages,
                max_tokens: 256,
                temperature: 0
            });

            const extracted = completion.choices?.[0]?.message?.content?.trim() ?? '';
            if (!extracted) return [query];

            try {
                const parsed = JSON.parse(extracted);
                if (Array.isArray(parsed)) return parsed.map((s: any) => String(s).trim()).filter(Boolean);
            } catch (e) {
                // Fallback: split on common delimiters
                const fallback = extracted.split(/[\n,;|]/).map(s => s.trim()).filter(Boolean);
                if (fallback.length) return fallback;
            }
        } catch (err) {
            console.error('ERROR: Azure OpenAI extraction failed', err);
        }

        // Default to using the raw query when extraction fails
        return [query];
    }

    // Run a SharePoint search for a single skill and return the list of display names.
    private async runSearchFor(skill: string): Promise<{ displayName: string; workEmail?: string }[]> {
        const siteUrl = "https://koskila.sharepoint.com";
        const sourceId = 'b09a7990-05ea-4af9-81ef-edfab16c4e31';

        const headersBase: Record<string, string> = {
            'Accept': 'application/json',
            'odata': 'verbose',
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + this._token
        };

        const safeSkill = skill.replace(/'/g, "''");
        const searchUrl = `${siteUrl}/_api/search/query?querytext='${encodeURIComponent(safeSkill)}'&sourceid='${sourceId}'`;

        try {
            const resp = await fetch(searchUrl, { method: 'GET', headers: headersBase });
            if (!resp.ok) {
                console.error(`ERROR: SharePoint search request failed for skill='${skill}'`, resp.status, await resp.text());
                return [];
            }

            const result = await resp.json();
            const rows = result?.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? [];

            const users = (rows as any[]).map((row: any) => {
                const cells = row?.Cells?.results ?? row?.Cells ?? [];
                const findValue = (keys: string[]) => {
                    for (const k of keys) {
                        const cell = (cells as any[]).find((c: any) => c?.Key === k);
                        if (cell && cell.Value) return cell.Value;
                    }
                    return null;
                };

                const name = findValue(['PreferredName', 'Title', 'AccountName']) ?? '';
                const email = findValue(['WorkEmail', 'AccountName', 'SPS-Mail']) ?? undefined;
                return { displayName: String(name).trim(), workEmail: email ? String(email).trim() : undefined };
            }).filter((u: any) => u.displayName) as { displayName: string; workEmail?: string }[];

            return users;
        } catch (err) {
            console.error(`ERROR: SharePoint search failed for skill='${skill}'`, err);
            return [];
        }
    }

    // Combine arrays of names into a single deduped array, preserving order.
    private combineResults(resultsArrays: { displayName: string; workEmail?: string }[][]): { displayName: string; workEmail?: string }[] {
        const seen = new Set<string>();
        const combined: { displayName: string; workEmail?: string }[] = [];
        for (const arr of resultsArrays) {
            for (const user of arr) {
                // Prefer dedupe by workEmail when available, otherwise by displayName
                const key = user.workEmail ? `email:${user.workEmail.toLowerCase()}` : `name:${user.displayName}`;
                if (!seen.has(key)) {
                    seen.add(key);
                    combined.push(user);
                }
            }
        }
        return combined;
    }
}
