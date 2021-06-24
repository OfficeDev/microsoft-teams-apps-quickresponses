// <copyright file="resources-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";

const baseAxiosUrl = window.location.origin;

/**
* Get resource data from API
* @param  {String | Null} token Custom jwt token
* @param  {String | Null} locale Current user selected application locale
*/
export const getResourceStrings = async (token: any, locale?: string | null): Promise<any> => {

	let url = baseAxiosUrl + "/api/Resource/ResourceStrings";
	return await axios.get(url, token, locale);
}

/**
* Get error resource data from API
* @param  {String | Null} token Custom jwt token
* @param  {String | Null} locale Current user selected application locale
*/
export const getErrorResourceStrings = async (token: any, locale?: string | null): Promise<any> => {

	let url = baseAxiosUrl + "/api/Resource/ErrorResourceStrings";
	return await axios.get(url, token, locale);
}