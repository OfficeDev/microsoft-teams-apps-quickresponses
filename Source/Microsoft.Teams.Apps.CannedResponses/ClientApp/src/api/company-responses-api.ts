// <copyright file="company-responses.api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";

const baseAxiosUrl = window.location.origin;

/**
* Get company responses data from API
* @param  {String | Null} token Custom jwt token
*/
export const getCompanyResponses = async (token: any): Promise<any> => {

	let url = baseAxiosUrl + "/api/companyresponse/GetCompanyResponses";
	return await axios.get(url, token);
}

/**
* Get user requests data from API
* @param  {String | Null} token Custom jwt token
*/
export const getUserRequests = async (token: any): Promise<any> => {

	let url = baseAxiosUrl + "/api/companyresponse/GetUserRequests";
	return await axios.get(url, token);
}

/**
* Delete user selected responses
*/
export const deleteUserResponseDetails = async (data: any[], token: any): Promise<any> => {

	let url = baseAxiosUrl + "/api/userresponse/deleteuserresponses";
	return await axios.post(url, data, token);
}