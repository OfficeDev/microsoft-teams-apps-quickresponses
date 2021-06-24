// <copyright file="response-message.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { useState, useEffect} from "react";
import { Icon, Text, Button } from "@fluentui/react";
import * as microsoftTeams from "@microsoft/teams-js";
import { createBrowserHistory } from "history";
import "../styles/site.css";
import { getApplicationInsightsInstance } from "../helpers/app-insights";
import { getResourceStringsFromApi } from "../helpers/resource-data";

const ResponseMessage: React.FunctionComponent<{}> = props => {
	let search = window.location.search;
	let params = new URLSearchParams(search);
	let theme = params.get("theme");
	let telemetry = params.get("telemetry");
	let locale = params.get("locale");
	let requestStatus = params.get("status");
	let messageText = params.get("message");
	let token = params.get("token");
	let isCompanyResponse = params.get("isCompanyResponse");
	const browserHistory = createBrowserHistory({ basename: "" });
	let appInsights;
	let userObjectId;
	let data;
	const [backButtonText, setBackButtonText] = useState("");

	useEffect(() => {
		microsoftTeams.initialize();
		microsoftTeams.getContext(async (context) => {
			userObjectId = context.userObjectId;
			locale = context.locale;
			appInsights = getApplicationInsightsInstance(telemetry, browserHistory);
			getResourceStrings();
		});
	}, []);

	/**
    *Get localized resource strings from API
    */
	async function getResourceStrings(){
		data = await getResourceStringsFromApi(appInsights, userObjectId, token, locale);
		if (data) {
			setBackButtonText(data.backButtonText);
		}
	}

	/**
	*Sets icons according to add and update request's response status
	*/
	function getIconComponent() {
		if (requestStatus === "addSuccess" || requestStatus === "editSuccess") {
			return (<Icon color="green" name="presence-available" className="response-message-icon" />);
		}
		else {
			return (<Icon color="red" name="error" className="response-message-icon" />);
		}
	}

	/**
	*Sets message according to add and update request's response status
	*/
	function getMessageText() {
		if (requestStatus === "addSuccess" || requestStatus === "editSuccess") {
			return (<div>
				<Text content={messageText} className="result-success-message-text" success size="largest" />
			</div>);
		}
		else {
			return (<div>
				<Text content={messageText} className="result-error-message-text" error size="largest" />
			</div>);
		}
	}

	/**
    * Triggers when user clicks back button
    */
	function onBackButtonClick() {
		if (isCompanyResponse==="true") {
			window.location.href = `/company-responses?theme=${theme}&token=${token}&telemetry=${telemetry}&locale=${locale}`;
		}
		else {
			window.location.href = `/user-responses?theme=${theme}&token=${token}&telemetry=${telemetry}&locale=${locale}`;
        }
	}

	return (
		<div>
			<div className="result-message-container">
				<div className="result-message-icon">
					{getIconComponent()}
				</div>

				<div className="result-message-text">
					{getMessageText()}
				</div>
			</div>
			<div className="add-form-button-container">
				<div>
					<Button icon="icon-chevron-start" content={backButtonText} text onClick={() => { onBackButtonClick() }} />
				</div>
			</div>
		</div>
	);
}

export default ResponseMessage;