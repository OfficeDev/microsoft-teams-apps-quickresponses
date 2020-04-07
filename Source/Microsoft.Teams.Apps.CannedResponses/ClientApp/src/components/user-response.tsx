// <copyright file="user-response.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import Resources from "../constants/resources";
import { Text, Loader } from "@fluentui/react";
import { ApplicationInsights, SeverityLevel } from "@microsoft/applicationinsights-web";
import * as microsoftTeams from "@microsoft/teams-js";
import { createBrowserHistory } from "history";
import CommandBar from "./user-response-command-bar";
import ResponsesTable from "./responses-table";
import { getResourceStringsFromApi } from "../helpers/resource-data";
import { getUserResponseDetails, deleteUserResponseDetails } from "../api/user-responses-api";
import { getApplicationInsightsInstance } from "../helpers/app-insights";
import "../styles/site.css";

const browserHistory = createBrowserHistory({ basename: "" });

interface IUserResponseData {
    userId: string,
    responseId: string,
    questionLabel: string,
    questionText: string,
    responseText: string,
    lastUpdatedDate: Date
}

interface IUserResponsesState {
    loader: boolean;
    theme: string;
    userResponsesData: IUserResponseData[];
    userSelectedResponses: string[];
    filteredUserResponses: IUserResponseData[];
    resourceStrings: any;
}

export interface IAppSettings {
    token: string | null,
    telemetry: string | null, 
    theme: string | null
}

export class UserResponsePage extends React.Component<{}, IUserResponsesState> {
    token?: string | null = null;
    telemetry?: any = null;
    theme?: string | null;
    locale?: string | null;
    userObjectId?: string = "";
    appInsights: ApplicationInsights;
    appSettings: IAppSettings = { telemetry: "", theme: "",token: "" };

    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.theme = this.appSettings.theme = params.get("theme");
        this.locale = params.get("locale");
        this.token = this.appSettings.token = params.get("token");
        this.telemetry = this.appSettings.telemetry = params.get("telemetry");

        this.state = {
            loader: true,
            theme: this.theme ? this.theme : Resources.default,
            filteredUserResponses: [],
            userResponsesData: [],
            userSelectedResponses: [],
            resourceStrings: {}
        }

        window.localStorage.setItem("appsettings", JSON.stringify(this.appSettings));

        // Initialize application insights for logging events and errors.
        this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
    }

    /**
    * Used to initialize Microsoft Teams sdk
    */
    async componentDidMount() {
        this.getResourceStrings();
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            this.getUserResponses();
        });
    }

    /**
    *Get user responses from API
    */
    async getUserResponses() {
        this.appInsights.trackTrace({ message: `'getUserResponseDetails' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let response = await getUserResponseDetails(this.userObjectId, this.token);
        if (response.status === 200 && response.data) {
            this.appInsights.trackTrace({ message: `'getUserResponseDetails' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            this.setState({
                userResponsesData: response.data,
                filteredUserResponses: response.data
            });
        }
        else {
            this.appInsights.trackTrace({ message: `'getUserResponseDetails' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }
        this.setState({
            loader: false
        });
    }

    /**
    *Get localized resource strings from API
    */
    async getResourceStrings() {
        let data = await getResourceStringsFromApi(this.appInsights, this.userObjectId, this.token, this.locale);

        if (data) {
            this.setState({
                resourceStrings: data
            });
        }
        this.setState({
            loader: false
        });
    }

    /**
    *Filters table as per search text entered by user
    *@param {String} searchText Search text entered by user
    */
    handleSearch = (searchText: string) => {
        if (searchText) {
            var filteredResponses = this.state.userResponsesData.filter(function (userResponse) {
                return userResponse.questionLabel.toUpperCase().includes(searchText.toUpperCase()) ||
                    userResponse.questionText.toUpperCase().includes(searchText.toUpperCase()) ||
                    userResponse.responseText.toUpperCase().includes(searchText.toUpperCase());
            });
            this.setState({ filteredUserResponses: filteredResponses });
        }
        else {
            this.setState({ filteredUserResponses: this.state.userResponsesData });
        }
    }

    onUserResponseSelected = (responseId: string, isSelected: boolean) => {
        if (isSelected) {
            let userSelectedResponses = this.state.userSelectedResponses;
            userSelectedResponses.push(responseId);
            this.setState({
                userSelectedResponses: userSelectedResponses
            })
        }
        else {
            let filteredUserResponses = this.state.userSelectedResponses.filter((addedResponseId) => {
                return addedResponseId !== responseId;
            });
            this.setState({
                userSelectedResponses: filteredUserResponses
            })
        }
    }

    /**
    *Navigate to add new response page
    */
    handleAddButtonClick = () => {
        this.appInsights.trackEvent({ name: `Add user response` }, { User: this.userObjectId });
        window.location.href = `/add-new-response?token=${this.token}&theme=${this.state.theme}&isNewAllowed=${!(this.state.userResponsesData.length >= 200)}&telemetry=${this.telemetry}`;
    }

    /**
    *Navigate to edit response page
    */
    handleEditButtonClick = () => {
        this.appInsights.trackEvent({ name: `Edit user response` }, { User: this.userObjectId });
        window.location.href = `/edit-user-response?id=${this.state.userSelectedResponses[0]}&theme=${this.state.theme}&token=${this.token}&telemetry=${this.telemetry}`;
    }

    /**
    *Deletes selected user responses
    */
    handleDeleteButtonClick = async () => {
        this.appInsights.trackEvent({ name: `Delete user response` }, { User: this.userObjectId });
        this.appInsights.trackTrace({ message: `'deleteUserResponseDetails' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let deletionResult = await deleteUserResponseDetails(this.state.userSelectedResponses, this.token);

        if (deletionResult.status === 200 && deletionResult.data) {
            this.appInsights.trackTrace({ message: `'deleteUserResponseDetails' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            let userResponses = this.state.userResponsesData.filter((userResponse) => {
                return !this.state.userSelectedResponses.includes(userResponse.responseId);
            });

            this.setState({
                userResponsesData: userResponses,
                filteredUserResponses: userResponses,
                userSelectedResponses: []
            })
        }
        else {
            this.appInsights.trackTrace({ message: `'deleteUserResponseDetails' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }
    }

    /**
    * Renders the component
    */
    public render(): JSX.Element {
        return (
            <div>
                {this.getWrapperPage()}
            </div>
        );
    }

    /**
    *Get wrapper for page which acts as container for all child components
    */
    private getWrapperPage = () => {
        if (this.state.loader) {
            return (
                    <div className="loader">
                        <Loader />
                    </div>
            );
        } else {
            const isDeleteButtonEnabled = this.state.userSelectedResponses.length > 0;
            const isEditButtonEnabled = this.state.userSelectedResponses.length > 0 && this.state.userSelectedResponses.length < 2; // Enable delete button when only one response row is selected.

            return (
                <div className="user-responses-wrapper-page">
                    <div>
                        <Text content={this.state.resourceStrings.myResponsesText} size="medium" />
                    </div>
                    <CommandBar
                        isDeleteEnable={isDeleteButtonEnabled}
                        isEditEnable={isEditButtonEnabled}
                        onAddButtonClick={this.handleAddButtonClick}
                        onDeleteButtonClick={this.handleDeleteButtonClick}
                        onEditButtonClick={this.handleEditButtonClick}
                        handleTableFilter={this.handleSearch}
                        resoureStrings={this.state.resourceStrings}
                    />
                    <div>
                        <ResponsesTable showCheckbox={true} responsesData={this.state.filteredUserResponses} onCheckBoxChecked={this.onUserResponseSelected} resoureStrings={this.state.resourceStrings} />
                    </div>
                </div>
            );
        }
    }
}
