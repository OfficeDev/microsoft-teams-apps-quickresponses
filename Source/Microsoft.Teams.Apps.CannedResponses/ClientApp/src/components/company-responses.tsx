// <copyright file="company-responses.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import Resources from "../constants/resources";
import { themes, Loader, Provider, Menu } from "@fluentui/react";
import * as microsoftTeams from "@microsoft/teams-js";
import { ApplicationInsights, SeverityLevel } from "@microsoft/applicationinsights-web";
import { createBrowserHistory } from "history";
import CompanyResponsesCommandBar from "./company-responses-command-bar";
import AddNewSuggestion from "./add-new-suggestion";
import ResponsesTable from "./responses-table";
import UserRequestsTable from "./user-requests-table";
import { getCompanyResponses, getUserRequests } from "../api/company-responses-api";
import { getResourceStringsFromApi } from "../helpers/resource-data";
import { getApplicationInsightsInstance } from "../helpers/app-insights";

import "../styles/site.css";

const browserHistory = createBrowserHistory({ basename: "" });

export interface ICompanyResponsesState {
    loader: boolean;
    theme: string;
    userRequests: any[];
    companyResponsesData: any[];
    filteredUserRequests: any[];
    userSelectedResponses: string[];
    filteredCompanyResponses: any[];
    menuItems: any[];
    showNewSuggestionForm: boolean;
    selectedMenuItemIndex: number;
    resourceStrings: any;
}

export interface IAppSettings {
    token: string | null,
    telemetry: string | null,
    theme: string | null
}

export class CompanyResponsePage extends React.Component<{}, ICompanyResponsesState> {

    token?: string | null = null;
    telemetry?: any = null;
    theme?: string | null;
    locale?: string | null;
    userObjectId?: string = "";
    appInsights: ApplicationInsights;
    appSettings: IAppSettings = { telemetry: "", theme: "", token: "" };

    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.theme = this.appSettings.theme = params.get("theme");
        this.token = this.appSettings.token = params.get("token");
        this.telemetry = this.appSettings.telemetry = params.get("telemetry");
        this.locale = params.get("locale");

        this.state = {
            loader: true,
            theme: this.theme ? this.theme : Resources.default,
            filteredCompanyResponses: [],
            companyResponsesData: [],
            userRequests: [],
            filteredUserRequests: [],
            userSelectedResponses: [],
            showNewSuggestionForm: false,
            selectedMenuItemIndex: 0,
            menuItems: [
                {
                    key: "1",
                    content: "Company responses",
                },
                {
                    key: "2",
                    content: "Requests",
                }
            ],
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
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            this.getResourceStrings();
            this.getCompanyRespose();
            this.getUserRequests();
        });
    }

    /**
    *Get company responses from API
    */
    async getCompanyRespose() {
        this.appInsights.trackTrace({ message: `'getCompanyResponses' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let response = await getCompanyResponses(this.token);

        if (response.status === 200 && response.data) {
            this.appInsights.trackTrace({ message: `'getCompanyResponses' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            this.setState({
                companyResponsesData: response.data,
                filteredCompanyResponses: response.data
            })
        }
        else {
            this.appInsights.trackTrace({ message: `'getCompanyResponses' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }
        this.setState({ loader: false });
    }

    /**
    *Get user suggestion for company responses from API
    */
    async getUserRequests() {
        this.appInsights.trackTrace({ message: `'getUserRequests' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let userRequestsResponse = await getUserRequests(this.token);

        if (userRequestsResponse.status === 200 && userRequestsResponse.data) {
            this.appInsights.trackTrace({ message: `'getUserRequests' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            this.setState({
                userRequests: userRequestsResponse.data,
                filteredUserRequests: userRequestsResponse.data
            })
        }
        else {
            this.appInsights.trackTrace({ message: `'getUserRequests' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }

        this.setState({ loader: false });
    }

    /**
    *Get localized resource strings from API
    */
    async getResourceStrings() {
        let data = await getResourceStringsFromApi(this.appInsights, this.userObjectId, this.token, this.locale);

        if (data) {
            var menuItems = [
                {
                    key: "1",
                    content: data.companyResponsesMenuText,
                },
                {
                    key: "2",
                    content: data.requestsMenuText,
                }
            ]
            this.setState({
                resourceStrings: data,
                menuItems: menuItems
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
            if (this.state.selectedMenuItemIndex === 0) {
                var filteredResponses = this.state.companyResponsesData.filter(function (userResponse) {
                    return userResponse.questionLabel.toUpperCase().includes(searchText.toUpperCase()) ||
                        userResponse.questionText.toUpperCase().includes(searchText.toUpperCase()) ||
                        userResponse.responseText.toUpperCase().includes(searchText.toUpperCase());
                });
                this.setState({ filteredCompanyResponses: filteredResponses });
            }
            else {
                var filteredResponses = this.state.userRequests.filter(function (userResponse) {
                    return userResponse.questionLabel.toUpperCase().includes(searchText.toUpperCase()) ||
                        userResponse.questionText.toUpperCase().includes(searchText.toUpperCase()) ||
                        userResponse.responseText.toUpperCase().includes(searchText.toUpperCase());
                });
                this.setState({ filteredUserRequests: filteredResponses });
            }
        }
        else {
            if (this.state.selectedMenuItemIndex === 0) {
                this.setState({ filteredCompanyResponses: this.state.companyResponsesData });
            }
            else {
                this.setState({ filteredUserRequests: this.state.userRequests });
            }
        }
    }

    /**
    * Triggers when user clicks back button
    */
    onBackButtonClick = () => {
        this.appInsights.trackEvent({ name: `Back` }, { User: this.userObjectId, FromPage:'company-responses' });
        this.setState({ showNewSuggestionForm: false });
    }

    /**
    *Display add new suggestion page
    */
    handleSuggestNewButtonClick = () => {
        this.appInsights.trackEvent({ name: `Suggest new company response` }, { User: this.userObjectId });
        this.setState({ showNewSuggestionForm: true });
    }

    /**
    * Renders the component
    */
    public render(): JSX.Element {
        const styleProps: any = {};
        switch (this.state.theme) {
            case Resources.dark:
                styleProps.theme = themes.teamsDark;
                break;
            case Resources.contrast:
                styleProps.theme = themes.teamsHighContrast;
                break;
            case Resources.light:
            default:
                styleProps.theme = themes.teams;
        }

        return (
            <div>
                {this.getWrapperPage(styleProps.theme)}
            </div>
        );
    }

    /** 
    *  Called once menu item is clicked.
    * */
    onMenuItemClick = async (event: any, data: any) => {
        this.setState({ selectedMenuItemIndex: data.index });
    }

    /**
    *Get wrapper for page which acts as container for all child components
    */
    private getWrapperPage = (theme: any) => {
        if (this.state.loader) {
            return (
                <Provider theme={theme}>
                    <div className="loader">
                        <Loader />
                    </div>
                </Provider>
            );
        } else {
            if (this.state.showNewSuggestionForm) {
                return <AddNewSuggestion isNewAllowed={!(this.state.companyResponsesData.length >= 1000)} resourceStrings={this.state.resourceStrings} onBackButtonClick={this.onBackButtonClick} />
            }
            else {
                return (
                    <div className="user-responses-wrapper-page">
                        <div>
                            <Menu defaultActiveIndex={0} onItemClick={this.onMenuItemClick} items={this.state.menuItems} styles = {{borderBottom:0}} className="menu" underlined primary />
                        </div>
                        <CompanyResponsesCommandBar
                            showSuggestNewButton={this.state.selectedMenuItemIndex === 0 ? true : false}
                            onSuggestNewButtonClick={this.handleSuggestNewButtonClick}
                            handleTableFilter={this.handleSearch}
                            resourceStrings={this.state.resourceStrings}
                        />
                        <div>
                            {this.state.selectedMenuItemIndex === 0 && <ResponsesTable resoureStrings={this.state.resourceStrings} showCheckbox={false} responsesData={this.state.filteredCompanyResponses} onCheckBoxChecked={() => { }} />}
                            {this.state.selectedMenuItemIndex === 1 && <UserRequestsTable resoureStrings={this.state.resourceStrings} requestsData={this.state.filteredUserRequests} />}
                        </div>
                    </div>
                );
            }
        }
    }
}