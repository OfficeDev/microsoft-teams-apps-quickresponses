// <copyright file="add-new-suggestion.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Flex, Text, Loader, Button, Input, TextArea } from "@fluentui/react";
import * as microsoftTeams from "@microsoft/teams-js";
import { ApplicationInsights } from "@microsoft/applicationinsights-web";
import { createBrowserHistory } from "history";
import { getApplicationInsightsInstance } from "../helpers/app-insights";

import "../styles/site.css";

const browserHistory = createBrowserHistory({ basename: "" });

export interface INewSuggestionState {
    loader: boolean;
    label: string;
    question: string;
    response: string;
    isLabelValuePresent: boolean;
    isQuestionValuePresent: boolean;
    isResponseValuePresent: boolean;
    isSubmitLoading: boolean;
}

interface INewSuggestionProps {
    onBackButtonClick: () => void,
    resourceStrings: any,
    isNewAllowed: boolean
}
export interface IAppSettings {
    token: string,
    telemetry: string
}

export default class AddNewSuggestion extends React.Component<INewSuggestionProps, INewSuggestionState> {

    token?: string | null = null;
    telemetry?: any = null;
    userObjectId?: string = "";
    appInsights: ApplicationInsights;
    upn?: string  = undefined;
    addNewSuggestionBotCommand:string = "AddNewSuggestion";

    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");

        this.state = {
            loader: false,
            label: "",
            question: "",
            response: "",
            isLabelValuePresent: true,
            isQuestionValuePresent: true,
            isResponseValuePresent: true,
            isSubmitLoading: false
        }

        // Initialize application insights for logging events and errors.
        this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
    }

    /**
    * Checks whether all validation conditions are matched before user submits new suggestion
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
    * Used to initialize Microsoft Teams sdk
    */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            this.upn = context.upn;
        });
    }

    /**
    * Submits new suggestion by user
    */
    onSuggestButtonClick = () => {
        if (this.checkIfSubmitAllowed()) {
            this.setState({ isSubmitLoading: true });
            this.appInsights.trackEvent({ name: `Suggest new` }, { User: this.userObjectId });
            let toBot = { Label: this.state.label, Question: this.state.question, Response: this.state.response, CommandContext: this.addNewSuggestionBotCommand, UPN: this.upn };
            microsoftTeams.tasks.submitTask(toBot);
        }
    }

    /**
    * Triggers when user clicks back button
    */
    onBackButtonClick = () => {
        this.appInsights.trackEvent({ name: `Back` }, { User: this.userObjectId, FromPage: 'add-new-suggestion' });
        this.props.onBackButtonClick();
    }

    /**
    * Set State value of category text box input control
    * @param {Any} event Object which describes occurred event
    */
    onLabelValueChange = (event: any) => {
        this.setState({ label: event.target.value, isLabelValuePresent: true  });
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
            return (
                <div className="add-user-responses-page">
                    <div className="add-form-container">
                        <div>
                            <Flex gap="gap.small">
                                <Text content={this.props.resourceStrings.label} size="medium" />
                                <Flex.Item push>
                                    {this.getUserResponsesLabelError(this.state.isLabelValuePresent)}
                                </Flex.Item>
                            </Flex>
                            <div className="add-form-input">
                                <Input placeholder={this.props.resourceStrings.typeCategoryPlaceholder} fluid required maxLength={200} value={this.state.label} onChange={this.onLabelValueChange} />
                            </div>
                        </div>
                        <div>
                            <Flex gap="gap.small">
                                <Text content={this.props.resourceStrings.questions} size="medium" />
                                <Flex.Item push>
                                    {this.getUserResponsesQuestionsError(this.state.isQuestionValuePresent)}
                                </Flex.Item>
                            </Flex>
                            <div className="add-form-input">
                                <Input placeholder={this.props.resourceStrings.typeQuestionPlaceholder} fluid required maxLength={500} value={this.state.question} onChange={this.onQuestionValueChange} />
                            </div>
                        </div>
                        <div>
                            <Flex gap="gap.small">
                                <Text content={this.props.resourceStrings.response} size="medium" />
                                <Flex.Item push>
                                    {this.getUserResponseError(this.state.isResponseValuePresent)}
                                </Flex.Item>
                            </Flex>
                            <div className="add-form-input">
                                <TextArea placeholder={this.props.resourceStrings.typeResponsePlaceholder} fluid required maxLength={500} className="response-text-area" value={this.state.response} onChange={this.onResponseValueChange} />
                            </div>
                        </div>
                    </div>
                    <div className="add-form-button-container">
                        <div>
                        <Flex space="between">
                            <Button icon="icon-chevron-start" content={this.props.resourceStrings.backButtonText} text onClick={() => { this.onBackButtonClick() }} />
                                <Flex gap="gap.small">
                                    <Button content={this.props.resourceStrings.suggestButtonText} loading={this.state.isSubmitLoading} disabled={this.state.isSubmitLoading || !this.props.isNewAllowed} primary onClick={() => { this.onSuggestButtonClick() }} />
                            </Flex>
                            </Flex>
                        </div>
                        <div>
                            {this.getUserResponsesMaximumAllowedError()}
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
            return (<Text content={this.props.resourceStrings.fieldRequiredMessage} className="field-error-message" error size="medium" />);
        }
        return (<></>);
    }

    /**
    *Returns text component containing error message for failed questions field validation
    *@param {boolean} isQuestionValuePresent Indicates whether questions value is present
    */
    private getUserResponsesQuestionsError = (isQuestionValuePresent: boolean) => {
        if (!isQuestionValuePresent) {
            return (<Text content={this.props.resourceStrings.fieldRequiredMessage} className="field-error-message" error size="medium" />);
        }
        return (<></>);
    }

    /**
    *Returns text component containing error message for failed response field validation
    *@param {boolean} isResponseValuePresent Indicates whether response value is present
    */
    private getUserResponseError = (isResponseValuePresent: boolean) => {
        if (!isResponseValuePresent) {
            return (<Text content={this.props.resourceStrings.fieldRequiredMessage} className="field-error-message" error size="medium" />);
        }
        return (<></>);
    }

    /**
    *Returns text component containing error message for failed maximum character validation
    *@param {boolean} isNewAllowed Indicates whether new response is allowed to add
    */
    private getUserResponsesMaximumAllowedError = () => {
        if (!this.props.isNewAllowed) {
            return (<Text content={this.props.resourceStrings.maxCompanyResponseMessage} className="max-error-message" error size="medium" />);
        }
    }
}