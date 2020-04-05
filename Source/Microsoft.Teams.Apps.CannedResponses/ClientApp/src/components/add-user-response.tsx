// <copyright file="add-user-response.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import Resources from "../constants/resources";
import { Flex, Text, Loader, Button, Input, TextArea } from "@fluentui/react";
import * as microsoftTeams from "@microsoft/teams-js";
import { ApplicationInsights } from "@microsoft/applicationinsights-web";
import { createBrowserHistory } from "history";
import { getResourceStringsFromApi } from "../helpers/resource-data";
import { getApplicationInsightsInstance } from "../helpers/app-insights";

import "../styles/site.css";

const browserHistory = createBrowserHistory({ basename: "" });

interface IAddUserResponsesState {
    loader: boolean;
    theme: string;
    label: string;
    question: string;
    response: string;
    isLabelValuePresent: boolean;
    isQuestionValuePresent: boolean;
    isResponseValuePresent: boolean;
    resourceStrings: any;
    isSubmitLoading: boolean;
}

export class AddUserResponse extends React.Component<{}, IAddUserResponsesState> {

    token?: string | null = null;
    telemetry?: any = null;
    theme?: string | null;
    isNewAllowed?: string | null;
    userObjectId?: string = "";
    locale?: string = "";
    appInsights: ApplicationInsights;
    addAndShareBotCommand: string = "AddAndShareUserResponse";
    addUserResponseBotCommand: string = "AddUserResponse";

    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.theme = params.get("theme");
        this.telemetry = params.get("telemetry");
        this.isNewAllowed = params.get("isNewAllowed");
        this.token = params.get("token");

        this.state = {
            loader: false,
            theme: this.theme ? this.theme : Resources.default,
            label: "",
            question: "",
            response: "",
            isLabelValuePresent: true,
            isQuestionValuePresent: true,
            isResponseValuePresent: true,
            resourceStrings: {},
            isSubmitLoading: false
        }

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
            this.locale = context.locale;
            this.getResourceStrings();
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
    *Submits and adds new user response
    */
    onAddButtonClick = () => {
        if (this.checkIfSubmitAllowed()) {
            this.setState({ isSubmitLoading: true });
            this.appInsights.trackEvent({ name: `Add user response` }, { User: this.userObjectId });
            let toBot = { Label: this.state.label, Question: this.state.question, Response: this.state.response, CommandContext: this.addUserResponseBotCommand };
            microsoftTeams.tasks.submitTask(toBot);
        }
    }

    /**
    *Checks whether all validation conditions are matched before user submits new response
    */
    checkIfSubmitAllowed = () => {
        if (this.state.label === "") {
            this.setState({ isLabelValuePresent: false });
        }

        if (this.state.question === "") {
            this.setState({ isQuestionValuePresent: false });
        }

        if (this.state.response === "") {
            this.setState({ isResponseValuePresent: false });
        }

        if (this.state.label && this.state.question && this.state.response) {
            return true;
        }
        else {
            return false;
        }
    }

    /**
    * Triggers when user clicks back button
    */
    onBackButtonClick = () => {
        this.appInsights.trackEvent({ name: `Back` }, { User: this.userObjectId, FromPage: 'add-user-response' });
        window.location.href = `/user-responses?theme=${this.state.theme}&token=${this.token}&telemetry=${this.telemetry}&locale=${this.locale}`;
    }

    /**
    * Set State value of category text box input control
    * @param {Any} event Object which describes occurred event
    */
    onLabelValueChange = (event: any) => {
        this.setState({ label: event.target.value, isLabelValuePresent: true });
    }

    /**
    * Set State value of questions text box input control
    *@param {Any} event Object which describes occurred event
    */
    onQuestionValueChange = (event: any) => {
        this.setState({ question: event.target.value, isQuestionValuePresent: true });
    }

    /**
    *Set State value of response text box input control
    *@param {Any} event Object which describes occurred event
    */
    onResponseValueChange = (event: any) => {
        this.setState({ response: event.target.value, isResponseValuePresent: true });
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
            let isNewAllowed: boolean;
            this.isNewAllowed === "false" ? isNewAllowed = false : isNewAllowed = true;

            return (
                <div className="add-user-responses-page">
                    <div>
                        <Text content={this.state.resourceStrings.myResponsesText} size="medium" />
                    </div>
                    <div className="add-form-container">
                        <div>
                            <Flex gap="gap.small">
                                <Text content={this.state.resourceStrings.label} size="medium" />
                                <Flex.Item push>
                                    {this.getUserResponsesLabelError(this.state.isLabelValuePresent)}
                                </Flex.Item>
                            </Flex>
                            <div className="add-form-input">
                                <Input placeholder={this.state.resourceStrings.typeCategoryPlaceholder} fluid required maxLength={200} value={this.state.label} onChange={this.onLabelValueChange} />
                            </div>
                        </div>
                        <div>
                            <Flex gap="gap.small">
                                <Text content={this.state.resourceStrings.questions} size="medium" />  
                                <Flex.Item push>
                                    {this.getUserResponsesQuestionsError(this.state.isQuestionValuePresent)}
                                </Flex.Item>
                            </Flex>
                            <div className="add-form-input">
                                <Input placeholder={this.state.resourceStrings.typeQuestionPlaceholder} fluid required maxLength={500} value={this.state.question} onChange={this.onQuestionValueChange} />
                            </div>
                        </div>
                        <div>
                            <Flex gap="gap.small">
                                <Text content={this.state.resourceStrings.response} size="medium" />
                                <Flex.Item push>
                                    {this.getUserResponseError(this.state.isResponseValuePresent)}
                                </Flex.Item>
                            </Flex>
                            <div className="add-form-input">
                                <TextArea placeholder={this.state.resourceStrings.typeResponsePlaceholder} fluid required maxLength={500} className="response-text-area" value={this.state.response} onChange={this.onResponseValueChange} />
                            </div>
                        </div>
                    </div>
                    <div className="add-form-button-container">
                        <div>
                            <Flex space="between">
                                <Button icon="icon-chevron-start" content={this.state.resourceStrings.backButtonText} text onClick={() => { this.onBackButtonClick() }} />
                                <Flex gap="gap.small">
                                    <Button content={this.state.resourceStrings.addButtonText} primary loading={this.state.isSubmitLoading} disabled={this.state.isSubmitLoading || !isNewAllowed} onClick={() => { this.onAddButtonClick() }} />
                                </Flex>
                            </Flex>
                        </div>
                        <div>
                            {this.getUserResponsesMaximumAllowedError(isNewAllowed)}
                        </div>
                    </div>
                </div>
            );
        }
    }

    /**
    *Returns text component containing error message for failed category field validation
    *@param {boolean} isLabelValuePresent Indicates whether category value is present
    */
    private getUserResponsesLabelError = (isLabelValuePresent: boolean) => {
        if (!isLabelValuePresent) {
            return (<Text content={this.state.resourceStrings.fieldRequiredMessage} className="field-error-message" error size="medium" />);
        }
        return (<></>);
    }

    /**
    *Returns text component containing error message for failed questions field validation
    *@param {boolean} isQuestionValuePresent Indicates whether questions value is present
    */
    private getUserResponsesQuestionsError = (isQuestionValuePresent: boolean) => {
        if (!isQuestionValuePresent) {
            return (<Text content={this.state.resourceStrings.fieldRequiredMessage} className="field-error-message" error size="medium" />);
        }
        return (<></>);
    }

    /**
    *Returns text component containing error message for failed response field validation
    *@param {boolean} isResponseValuePresent Indicates whether response value is present
    */
    private getUserResponseError = (isResponseValuePresent: boolean) => {
        if (!isResponseValuePresent) {
            return (<Text content={this.state.resourceStrings.fieldRequiredMessage} className="field-error-message" error size="medium" />);
        }
        return (<></>);
    }

    /**
    *Returns text component containing error message for failed maximum character validation
    *@param {boolean} isNewAllowed Indicates whether new response is allowed to add
    */
    private getUserResponsesMaximumAllowedError = (isNewAllowed: boolean) => {
        if (!isNewAllowed) {
            return (<Text content={this.state.resourceStrings.maxResponsesMessage} className="max-error-message" error size="medium" />);
        }
    }
}