// <copyright file="error-page.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { RouteComponentProps } from "react-router-dom";
import { Text, Loader } from "@fluentui/react";
import { IAppSettings } from "./user-response";
import { getErrorResourceStringsFromApi } from "../helpers/resource-data";

import "../styles/site.css";

interface IResourceString {
    unauthorizedErrorMessage: string,
    forbiddenErrorMessage: string,
    generalErrorMessage: string,
    refreshLinkText: string
}

interface errorPageState {
    loader: boolean;
    resourceStrings: IResourceString,
}

export class ErrorPage extends React.Component<RouteComponentProps, errorPageState> {   
    locale: string = "";
    private appSettings: IAppSettings = {
        telemetry: "",
        token: "",
        theme: ""
    };

    constructor(props: any) {
        super(props);
        this.state = {
            loader: true,
            resourceStrings: {
                unauthorizedErrorMessage: "Sorry, an error occurred while trying to access this service.",
                forbiddenErrorMessage: "Sorry, seems like you don't have permission to access this page.",
                generalErrorMessage: "Oops! An unexpected error seems to have occured. Why not try refreshing your page? Or you can contact your administrator if the problem persists.",
                refreshLinkText: "Refresh"
            }
        };
        let storageValue = localStorage.getItem("appsettings");
        if (storageValue) {
            this.appSettings = JSON.parse(storageValue) as IAppSettings;
        }
    }

    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.locale = context.locale;
            this.getResourceStrings();
        });
    }

    /**
    *Get localized resource strings from API
    */
    async getResourceStrings() {

        let data = await getErrorResourceStringsFromApi(this.appSettings.token, this.locale);
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
    * Renders the component
    */
    public render(): JSX.Element {
        const params = this.props.match.params;
        let message = `${this.state.resourceStrings.generalErrorMessage}`;

        if ("id" in params) {
            const id = params["id"];
            if (id === "401") {
                message = `${this.state.resourceStrings.unauthorizedErrorMessage}`;
            } else if (id === "403") {
                message = `${this.state.resourceStrings.forbiddenErrorMessage}`;
            }
            else {
                message = `${this.state.resourceStrings.generalErrorMessage}`;
            }
        }
        if (!this.state.loader) {
            return (
                <div className="error-message">
                    <Text content={message}  error size="medium" />
                </div>
            );
        }
        else {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        }
    }
}