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

    public async getUserBySkills(query: string): Promise<{ displayName: string; workEmail?: string; skills?: string[]; department?: string; location?: string }[]> {
        // Extract filters (skills, department, office/location) from the query,
        // build one or more SharePoint search queries that support any combination
        // of Skills, Department and OfficeNumber, then combine and dedupe results.
        const filters = await this.extractFilters(query);
        const uniqueSkills = Array.from(new Set((filters.skills || []).map(s => s.trim()).filter(Boolean)));

        const searchQueries: string[] = [];

        // Helper to quote values that contain spaces or special chars
        const q = (v: string) => (/[\s:\(\)\"]/.test(v) ? `"${v}"` : v);

        if (uniqueSkills.length > 0) {
            // For each skill, create a query that optionally includes Department/OfficeNumber filters
            for (const skill of uniqueSkills) {
                let part = `Skills:${q(skill)}`;
                if (filters.department) part += ` AND Department:${q(filters.department)}`;
                if (filters.officeNumber) part += ` AND OfficeNumber:${q(filters.officeNumber)}`;
                searchQueries.push(`(${part})`);
            }
        } else if (filters.department || filters.officeNumber) {
            // No skills, but department/office present -> single combined query
            const parts: string[] = [];
            if (filters.officeNumber) parts.push(`OfficeNumber:${q(filters.officeNumber)}`);
            if (filters.department) parts.push(`Department:${q(filters.department)}`);
            searchQueries.push(`(${parts.join(' AND ')})`);
        } else {
            // Fallback: search by the raw query string in Skills
            searchQueries.push(`(Skills:${q(query)})`);
        }

        // Run searches in parallel and return the combined, deduped results
        const resultsArrays = await Promise.all(searchQueries.map(s => this.runSearchFor(s)));
        return this.combineResults(resultsArrays);
    }

    // Extract skills, department and office/location from the user's query using Azure OpenAI when configured.
    // Returns an object: { skills: string[], department?: string, officeNumber?: string }
    private async extractFilters(query: string): Promise<{ skills: string[]; department?: string; officeNumber?: string }> {
        const openaiEndpoint = process.env.AZURE_OPENAI_ENDPOINT;
        const openaiApiKey = process.env.AZURE_OPENAI_API_KEY;
        const openaiDeployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;

        // Default fallback
        const defaultResult = { skills: [query] };

        if (!openaiEndpoint || !openaiApiKey) {
            console.warn('AZURE_OPENAI_ENDPOINT or AZURE_OPENAI_API_KEY not set; skipping extraction and using provided query');
            return defaultResult;
        }

        try {
            const apiVersion = process.env.OPENAI_API_VERSION || '2024-10-21';
            const client = new AzureOpenAI({ endpoint: openaiEndpoint, apiKey: openaiApiKey, deployment: openaiDeployment, apiVersion });

            const messages: ChatCompletionMessageParam[] = [
                { role: 'system', content: 'You are a strict extractor. When asked, return ONLY a JSON object with the keys: "skills" (an array of skill strings), "department" (string) and "officeNumber" (string). Example: {"skills":["Strategic Thinking","Team Building"], "department":"Quality Assurance", "officeNumber":"Helsinki"}. If a value is not present, omit the key or use an empty array for skills. Return NO extra text.' },
                { role: 'user', content: `Extract skills, department and office/location from the following prompt. Return ONLY a JSON object as described above. Prompt:\n\n${query}` }
            ];

            const completion = await client.chat.completions.create({
                model: openaiDeployment!,
                messages: messages,
                max_tokens: 256,
                temperature: 0
            });

            const extracted = completion.choices?.[0]?.message?.content?.trim() ?? '';
            if (!extracted) return defaultResult;

            try {
                const parsed = JSON.parse(extracted);
                if (parsed && typeof parsed === 'object') {
                    const skills = Array.isArray(parsed.skills) ? parsed.skills.map((s: any) => String(s).trim()).filter(Boolean) : [];
                    // accept common alternative keys
                    const department = parsed.department ?? parsed.Department ?? parsed.dept ?? parsed.Dept ?? parsed.departmentName ?? undefined;
                    const officeNumber = parsed.officeNumber ?? parsed.office ?? parsed.Office ?? parsed.location ?? parsed.Location ?? undefined;
                    const result: { skills: string[]; department?: string; officeNumber?: string } = { skills };
                    if (department && String(department).trim()) result.department = String(department).trim();
                    if (officeNumber && String(officeNumber).trim()) result.officeNumber = String(officeNumber).trim();
                    if (result.skills.length) return result;
                    // If no skills but department/office present, still return
                    if (result.department || result.officeNumber) return result;
                }
            } catch (e) {
                // fall through to heuristic fallback below
            }

            // Fallback: try simple parsing from the extracted text or original query
            // Try to split on common delimiters for skills
            const fallbackSkills = extracted.split(/[\n,;|\/]+/).map(s => s.trim()).filter(Boolean);
            // Try to detect department/office patterns in the extracted text first, then in the original query
            const detect = (text: string) => {
                // explicit labels
                const deptMatch = /Department\s*[:\-]\s*([^,;\n\)\(]+)/i.exec(text)
                    || /Dept\.?\s*[:\-]\s*([^,;\n\)\(]+)/i.exec(text)
                    || /([A-Za-z &\-]+)\s+department/i.exec(text)
                    || /department\s+of\s+([^,;\n\)\(]+)/i.exec(text);

                const officeMatch = /Office(?:Number)?\s*[:\-]\s*([^,;\n\)\(]+)/i.exec(text)
                    || /Location\s*[:\-]\s*([^,;\n\)\(]+)/i.exec(text)
                    || /based in\s+([A-Z][A-Za-z\-\s&]+)/i.exec(text)
                    || /located in\s+([A-Z][A-Za-z\-\s&]+)/i.exec(text)
                    || /\b(?:in|at)\s+([A-Z][A-Za-z\-\s&]+)/i.exec(text)
                    || /office in\s+([A-Z][A-Za-z\-\s&]+)/i.exec(text);

                // prefer shorter captures trimmed to first token if long
                const clean = (s?: string) => {
                    if (!s) return undefined;
                    const v = s.trim();
                    // stop at comma/semicolon/newline if present
                    return v.split(/[,;\n]/)[0].trim();
                };

                return { department: clean(deptMatch?.[1]), officeNumber: clean(officeMatch?.[1]) };
            };

            let heur = detect(extracted);
            if (!heur.department && !heur.officeNumber) heur = detect(query);

            const result: { skills: string[]; department?: string; officeNumber?: string } = { skills: fallbackSkills.length ? fallbackSkills : [query] };
            if (heur.department) result.department = heur.department;
            if (heur.officeNumber) result.officeNumber = heur.officeNumber;
            return result;
        } catch (err) {
            console.error('ERROR: Azure OpenAI extraction failed', err);
        }

        // Default to using the raw query when extraction fails
        return defaultResult;
    }

    // Run a SharePoint search for a single search query and return the list of display names.
    // The input is a raw querystring that can include fielded operators such as
    // (Skills: Leadership AND Department:"Quality Assurance") or (OfficeNumber:Helsinki)
    private async runSearchFor(searchQuery: string): Promise<{ displayName: string; workEmail?: string; skills?: string[]; department?: string; location?: string }[]> {
        const siteUrl = "https://koskila.sharepoint.com";
        const sourceId = 'b09a7990-05ea-4af9-81ef-edfab16c4e31';

        const headersBase: Record<string, string> = {
            'Accept': 'application/json',            
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + this._token
        };

    const safeQuery = searchQuery.replace(/'/g, "''");
    const searchUrl = `${siteUrl}/_api/search/query?querytext='${encodeURIComponent(safeQuery)}'&sourceid='${sourceId}'&selectproperties='PreferredName,WorkEmail,Skills,Department,Location,Title,SPS-Mail,Tags,Department,OfficeNumber,BaseOfficeLocation,SPS-Department,Office,AccountName,PeopleKeywords'`;

        try {
            const resp = await fetch(searchUrl, { method: 'GET', headers: headersBase });
            if (!resp.ok) {
                console.error(`ERROR: SharePoint search request failed for query='${searchQuery}'`, resp.status, await resp.text());
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
                // Try to find skills, department and location from common SharePoint managed properties
                const rawSkills = findValue(['Skills', 'PeopleKeywords', 'Tags', 'RefinableString01']) ?? '';
                const department = (findValue(['Department', 'SPS-Department', 'Office']) ?? undefined) as string | undefined;
                const location = (findValue(['Office', 'SPS-Location', 'Location', 'OfficeNumber']) ?? undefined) as string | undefined;

                // Normalize skills into an array (split on common delimiters)
                const skills: string[] = String(rawSkills)
                    .split(/[,;|\n\/]+/)
                    .map((s) => s.trim())
                    .filter(Boolean);

                return {
                    displayName: String(name).trim(),
                    workEmail: email ? String(email).trim() : undefined,
                    skills: skills.length ? skills : undefined,
                    department: department ? String(department).trim() : undefined,
                    location: location ? String(location).trim() : undefined
                };
            }).filter((u: any) => u.displayName) as { displayName: string; workEmail?: string }[];

            return users;
        } catch (err) {
            console.error(`ERROR: SharePoint search failed for query='${searchQuery}'`, err);
            return [];
        }
    }

    // Combine arrays of users into a single deduped array, preserving order.
    // When duplicates are found we merge skills arrays and prefer non-empty department/location values.
    private combineResults(resultsArrays: { displayName: string; workEmail?: string; skills?: string[]; department?: string; location?: string }[][]): { displayName: string; workEmail?: string; skills?: string[]; department?: string; location?: string }[] {
        const seen = new Set<string>();
        const combined: { displayName: string; workEmail?: string; skills?: string[]; department?: string; location?: string }[] = [];
        for (const arr of resultsArrays) {
            for (const user of arr) {
                const key = user.workEmail ? `email:${user.workEmail.toLowerCase()}` : `name:${user.displayName}`;
                if (!seen.has(key)) {
                    seen.add(key);
                    // Clone to avoid accidental mutation
                    combined.push({
                        displayName: user.displayName,
                        workEmail: user.workEmail,
                        skills: user.skills ? Array.from(new Set(user.skills)) : undefined,
                        department: user.department,
                        location: user.location
                    });
                } else {
                    // Merge into existing entry: add skills, fill department/location when missing
                    const existing = combined.find((u) => (u.workEmail ? `email:${u.workEmail!.toLowerCase()}` : `name:${u.displayName}`) === key)!;
                    if (user.skills && user.skills.length) {
                        const mergedSkills = new Set([...(existing.skills || []), ...user.skills]);
                        existing.skills = Array.from(mergedSkills);
                    }
                    if (!existing.department && user.department) existing.department = user.department;
                    if (!existing.location && user.location) existing.location = user.location;
                }
            }
        }
        return combined;
    }
}
