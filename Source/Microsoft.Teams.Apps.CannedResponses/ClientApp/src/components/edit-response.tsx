// <copyright file="add-user-response.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import Resources from "../constants/resources";
import { Flex, Text, Loader, Button, Input, TextArea } from "@fluentui/react";
import * as microsoftTeams from "@microsoft/teams-js";
import { ApplicationInsights, SeverityLevel } from "@microsoft/applicationinsights-web";
import { createBrowserHistory } from "history";
import { getUserResponseDetailsForResponseId } from "../api/user-responses-api";
import { getResourceStringsFromApi } from "../helpers/resource-data";
import { getApplicationInsightsInstance } from "../helpers/app-insights";

import "../styles/site.css";

const browserHistory = createBrowserHistory({ basename: "" });

interface IEditUserResponsesState {
    loader: boolean;
    theme: string;
    label: string;
    question: string;
    response: string;
    responseData: any;
    isLabelValuePresent: boolean;
    isQuestionValuePresent: boolean;
    isResponseValuePresent: boolean;
    resourceStrings: any;
    isSubmitLoading: boolean;
}

export class EditUserResponse extends React.Component<{}, IEditUserResponsesState> {

    token?: string | null = null;
    responseId: string | null = null;
    telemetry?: any = null;
    theme?: string | null;
    locale?: string | null;
    userObjectId?: string = "";
    appInsights: ApplicationInsights;

    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.theme = params.get("theme");
        this.telemetry = params.get("telemetry");
        this.responseId = params.get("id");
        this.token = params.get("token");

        this.state = {
            loader: true,
            theme: this.theme ? this.theme : Resources.default,
            label: "",
            question: "",
            response: "",
            responseData: null,
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
        this.getResourceStrings();
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            this.locale = context.locale;
            this.getResourceStrings();
            this.getUserResponseDetailsForResponseId();
        }); 
    }

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

    async getUserResponseDetailsForResponseId() {
        this.appInsights.trackTrace({ message: `'getUserResponseDetailsForResponseId' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let response = await getUserResponseDetailsForResponseId(this.responseId, this.token);

        if (response.status === 200 && response.data) {
            this.appInsights.trackTrace({ message: `'getUserResponseDetailsForResponseId' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            this.setState({
                responseData: response.data,
                label: response.data.questionLabel,
                question: response.data.questionText,
                response: response.data.responseText,
            });
        }
        else {
            this.appInsights.trackTrace({ message: `'getUserResponseDetailsForResponseId' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }
        this.setState({
            loader: false
        });
    }

    onUpdateButtonClick = () => {
        if (this.checkIfSubmitAllowed()) {
            this.setState({ isSubmitLoading: true });
            this.appInsights.trackEvent({ name: `Update user response` }, { User: this.userObjectId });
            let toBot = { ResponseId: this.state.responseData.responseId, Label: this.state.label, Question: this.state.question, Response: this.state.response, CommandContext: "EditUserResponse" };
            microsoftTeams.tasks.submitTask(toBot);
        }
    }

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

    onBackButtonClick = () => {
        this.appInsights.trackEvent({ name: `Back` }, { User: this.userObjectId, FromPage: 'edit-response' });
        window.location.href = `/user-responses?theme=${this.state.theme}&token=${this.token}&telemetry=${this.telemetry}&locale=${this.locale}`;
    }

    onLabelValueChange = (event: any) => {
        let responseDataValue = this.state.responseData;
        responseDataValue.label = event.target.value
        this.setState({ responseData: responseDataValue, label: responseDataValue.label , isLabelValuePresent: true });
    }

    onQuestionValueChange = (event: any) => {
        let responseDataValue = this.state.responseData;
        responseDataValue.question = event.target.value
        this.setState({ responseData: responseDataValue, question: responseDataValue.question, isQuestionValuePresent: true });
    }

    onResponseValueChange = (event: any) => {
        let responseDataValue = this.state.responseData;
        responseDataValue.response = event.target.value
        this.setState({ responseData: responseDataValue, response: responseDataValue.response, isResponseValuePresent: true });
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

    private getWrapperPage = () => {
        if (!this.state.responseData) {
            return (
                    <div className="loader">
                        <Loader />
                    </div>
            );
        } else {
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
                                <Input placeholder={this.state.resourceStrings.typeQuestionPlaceholder} fluid required maxLength={500} value={this.state.question} onChange={this.onQuestionValueChange} /><br />
                                <Text content={this.state.resourceStrings.typeQuestionPlaceholder} size="small" styles={{ float: "right" }} />
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
                        <Flex space="between">
                            <Button icon="icon-chevron-start" content={this.state.resourceStrings.backButtonText} text onClick={() => { this.onBackButtonClick() }} />
                            <Flex gap="gap.small">
                                <Button content={this.state.resourceStrings.updateButtonText} loading={this.state.isSubmitLoading} disabled={this.state.isSubmitLoading} primary onClick={() => { this.onUpdateButtonClick() }} />
                            </Flex>
                        </Flex>
                    </div>
                </div>
            );
        }
    }

    private getUserResponsesLabelError = (isLabelValuePresent: boolean) => {
        if (!isLabelValuePresent) {
            return (<Text content={this.state.resourceStrings.fieldRequiredMessage} className="field-error-message" error size="medium" />);
        }
        return (<></>);
    }

    private getUserResponsesQuestionsError = (isQuestionValuePresent: boolean) => {
        if (!isQuestionValuePresent) {
            return (<Text content={this.state.resourceStrings.fieldRequiredMessage} className="field-error-message" error size="medium" />);
        }
        return (<></>);
    }

    private getUserResponseError = (isResponseValuePresent: boolean) => {
        if (!isResponseValuePresent) {
            return (<Text content={this.state.resourceStrings.fieldRequiredMessage} className="field-error-message" error size="medium" />);
        }
        return (<></>);
    }
}