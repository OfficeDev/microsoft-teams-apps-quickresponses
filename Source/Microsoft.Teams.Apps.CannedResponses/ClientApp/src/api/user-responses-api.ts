// <copyright file="user-responses-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";

const baseAxiosUrl = window.location.origin;

/**
* Get user responses data from API
* @param {String} userId Unique user object identifier for which user responses will be fetched
* @param {String | Null} token Custom JWT token
*/
export const getUserResponseDetails = async (userId: string | undefined, token: any): Promise<any> => {

	let url = baseAxiosUrl + `/api/userresponse?id${userId}`;
	return await axios.get(url, token);
}

/**
* Get user response details from API
* @param {String | Null} responseId Unique response ID for which details will be fetched
* @param {String | Null} token Custom JWT token
*/
export const getUserResponseDetailsForResponseId = async (responseId: string | null, token: any): Promise<any> => {

	let url = baseAxiosUrl + `/api/userresponse/responsedata?responseId=${responseId}`;
	return await axios.get(url, token);
}

/**
* Delete user selected responses
* @param {Array<any>} data Selected user responses which needs to be deleted
* @param {String | Null} token Custom JWT token
*/
export const deleteUserResponseDetails = async (data: any[], token: any): Promise<any> => {

	let url = baseAxiosUrl + "/api/userresponse/deleteresponses";
	return await axios.post(url, data, token);
}