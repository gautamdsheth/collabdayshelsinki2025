// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import { AzureOpenAI } from 'openai';
import dotenv from 'dotenv';

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

    public async getUserBySkills(query: string): Promise<string[]> {
        // First, attempt to extract an exact "leadership skills" value from the provided prompt
        // using Azure OpenAI Chat Completions. If extraction fails or env vars are missing,
        // fall back to using the provided query as-is.

        const openaiEndpoint = process.env.AZURE_OPENAI_ENDPOINT;
        const openaiApiKey = process.env.AZURE_OPENAI_API_KEY;
        const openaiDeployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;

        let searchProperty = query;

        if (openaiEndpoint && openaiApiKey) {
            try {
                const apiVersion = process.env.OPENAI_API_VERSION || '2024-10-21';
                const client = new AzureOpenAI({ endpoint: openaiEndpoint, apiKey: openaiApiKey, deployment: openaiDeployment, apiVersion });

                const messages = [
                    { role: 'system', content: 'You are a strict skills extractor. When asked for a value, return only the exact value with no extra text.' },
                    { role: 'user', content: `Extract the exact value for "leadership" from the following prompt. Return only the exact value (for example: "Strategic Thinking") or an empty string if none found. Prompt:\n\n${query}` }
                ];

                const completion = await client.chat.completions.create({
                    model: openaiDeployment,
                    messages: messages as any,
                    max_tokens: 128,
                    temperature: 0
                });

                const extracted = completion.choices?.[0]?.message?.content?.trim() ?? '';

                if (extracted) {
                    searchProperty = extracted.replace(/'/g, "''");
                }
            } catch (err) {
                console.error('ERROR: Azure OpenAI extraction failed', err);
                searchProperty = query;
            }
        } else {
            console.warn('AZURE_OPENAI_ENDPOINT or AZURE_OPENAI_API_KEY not set; skipping extraction and using provided query');
        }

        // Use SharePoint Search REST API instead of Microsoft Graph for this query.
        // Requires `SP_SITE_URL` to be set in the environment (e.g. https://contoso.sharepoint.com/sites/mysite)

        const siteUrl = "https://koskila.sharepoint.com";

        // Build search URL similar to the PowerShell snippet:
        // $urlDefaultSite/_api/search/query?querytext='<searchProperty>:true'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'
        const sourceId = 'b09a7990-05ea-4af9-81ef-edfab16c4e31';
        const searchUrl = `${siteUrl}/_api/search/query?querytext='${encodeURIComponent(searchProperty)}'&sourceid='${sourceId}'`;

        const headers: Record<string, string> = {
            'Accept': 'application/json',
            'odata': 'verbose',
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + this._token
        };

        const response = await fetch(searchUrl, { method: 'GET', headers });
        if (!response.ok) {
            // Log error and return empty list to preserve original method's contract
            console.error('ERROR: SharePoint search request failed', response.status, await response.text());
            return [];
        }

        const result = await response.json();

        // The SharePoint search response places rows under PrimaryQueryResult.RelevantResults.Table.Rows
        const rows = result?.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? [];

        // Each row contains Cells (array) with Key/Value pairs. Try to extract a display name from common keys.
        const displayNames = (rows as any[]).map((row: any) => {
            const cells = row?.Cells?.results ?? row?.Cells ?? [];
            const findValue = (keys: string[]) => {
                for (const k of keys) {
                    const cell = (cells as any[]).find((c: any) => c?.Key === k);
                    if (cell && cell.Value) return cell.Value;
                }
                return null;
            };

            // Common keys that may contain the user's display name
            return findValue(['PreferredName', 'Title', 'AccountName', 'WorkEmail']) ?? '';
        }).filter(Boolean);

        return displayNames;
    }    

    // Gets the user's photo
    public async getProfilePhotoAsync(profile: any): Promise<string> {
        const graphPhotoEndpoint = `https://graph.microsoft.com/v1.0/users/${profile.id}/photos/240x240/$value`;
        const graphRequestParams = {
            method: 'GET',
            headers: {
                'Content-Type': 'image/png',
                authorization: 'bearer ' + this._token
            }
        };

        const response = await fetch(graphPhotoEndpoint, graphRequestParams);
        if (!response.ok) {
            console.error('ERROR: ', response);
        }

        const imageBuffer = await response.arrayBuffer(); //Get image data as raw binary data

        //Convert binary data to an image URL and set the url in state
        const imageUri = 'data:image/png;base64,' + Buffer.from(imageBuffer).toString('base64');
        return imageUri;
    }
}
