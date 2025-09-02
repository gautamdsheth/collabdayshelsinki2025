// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { Attachment, CardFactory, CardAction, MessagingExtensionAttachment } from 'botbuilder';

/**
 * Creates an adaptive card for an npm package search result.
 * @param {any} result The search result to create the card for.
 * @returns {Attachment} The adaptive card attachment for the search result.
 */
export function createNpmPackageCard(result: any): Attachment {
    // Simplified card: only show maintainers with display name and email
    const body: any[] = [
        {
            type: 'TextBlock',
            text: 'Maintainers',
            weight: 'Bolder',
            size: 'Medium'
        }
    ];

    if (result.maintainers && result.maintainers.length) {
        for (const m of result.maintainers) {
            body.push({
                type: 'ColumnSet',
                columns: [
                    {
                        type: 'Column',
                        width: 'stretch',
                        items: [
                            {
                                type: 'TextBlock',
                                text: m.name ?? m.email ?? 'Unknown',
                                wrap: true
                            }
                        ]
                    },
                    {
                        type: 'Column',
                        width: 'auto',
                        items: [
                            {
                                type: 'TextBlock',
                                text: m.email ?? '',
                                isSubtle: true,
                                wrap: true,
                                ...(m.email ? { selectAction: { type: 'Action.OpenUrl', url: `mailto:${m.email}` } } : {})
                            }
                        ]
                    }
                ]
            });
        }
    } else {
        body.push({
            type: 'TextBlock',
            text: 'No maintainers found',
            isSubtle: true
        });
    }

    return CardFactory.adaptiveCard({
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        type: 'AdaptiveCard',
        version: '1.2',
        body
    });
}

/**
 * Creates a messaging extension attachment for an npm search result.
 * @param {any} result The search result to create the attachment for.
 * @returns {MessagingExtensionAttachment} The messaging extension attachment for the search result.
 */
export function createNpmSearchResultCard(result: any): MessagingExtensionAttachment {
    const card = CardFactory.heroCard(result.name, [], [], {
        text: result.description
    }) as MessagingExtensionAttachment;
    card.preview = CardFactory.heroCard(result.name, [], [], {
        tap: { type: 'invoke', value: result } as CardAction,
        text: result.description
    });
    return card;
}

/**
 * @returns {Attachment} The adaptive card attachment for the sign-in request.
 */
export function createSignOutCard(): Attachment {
    return CardFactory.adaptiveCard({
        version: '1.0.0',
        type: 'AdaptiveCard',
        body: [
            {
                type: 'TextBlock',
                text: 'You have been signed out.'
            }
        ],
        actions: [
            {
                type: 'Action.Submit',
                title: 'Close',
                data: {
                    key: 'close'
                }
            }
        ]
    });
}

/**
 *
 * @param {string} displayName The display name of the user
 * @param {string} profilePhoto The profile photo of the user
 * @returns {Attachment} The adaptive card attachment for the user profile.
 */
export function createUserProfileCard(displayName: string, profilePhoto: string): Attachment {
    return CardFactory.adaptiveCard({
        version: '1.0.0',
        type: 'AdaptiveCard',
        body: [
            {
                type: 'TextBlock',
                text: 'Hello: ' + displayName
            },
            {
                type: 'Image',
                url: profilePhoto
            }
        ]
    });
}
