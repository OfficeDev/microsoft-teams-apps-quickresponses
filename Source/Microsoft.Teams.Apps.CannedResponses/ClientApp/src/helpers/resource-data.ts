// <copyright file="resource-data.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import { ApplicationInsights, SeverityLevel } from "@microsoft/applicationinsights-web";
import { getResourceStrings, getErrorResourceStrings } from "../api/resources-api";

export const getResourceStringsFromApi = async (appInsights: ApplicationInsights, userObjectId: string | undefined, token?: string | null, locale?: string | null): Promise<any> => {
	appInsights.trackTrace({ message: `'getResourceStrings' - Initiated request`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
	let response = await getResourceStrings(token, locale);
	if (response.status === 200 && response.data) {
		appInsights.trackTrace({ message: `'getResourceStrings' - Request success`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
		return response.data;
	}
	else {
		appInsights.trackTrace({ message: `'getResourceStrings' - Request failed`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
		return null;
	}
}

export const getErrorResourceStringsFromApi = async (token?: string | null): Promise<any> => {
	let response = await getErrorResourceStrings(token);
	if (response.status === 200 && response.data) {
		return response.data;
	}
	else {
		console.log("Error occurred while getting error resource strings from api.")
		return null;
	}
}