// <copyright file="app.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { AppRoute } from "./router/router";
import Resources from "./constants/resources";
import { Provider, themes } from "@fluentui/react";

export interface IAppState {
	theme: string;
}

export default class App extends React.Component<{}, IAppState> {
	theme?: string | null;

	constructor(props: any) {
		super(props);
		let search = window.location.search;
		let params = new URLSearchParams(search);
		this.theme = params.get("theme");

		this.state = {
			theme: this.theme ? this.theme : Resources.default,
		}
	}

	public setThemeComponent = () => {
		if (this.state.theme === "dark") {
			return (
				<Provider theme={themes.teamsDark}>
					<div className="darkContainer">
						{this.getAppDom()}
					</div>
				</Provider>
			);
		}
		else if (this.state.theme === "contrast") {
			return (
				<Provider theme={themes.teamsHighContrast}>
					<div className="highContrastContainer">
						{this.getAppDom()}
					</div>
				</Provider>
			);
		} else {
			return (
				<Provider theme={themes.teams}>
					<div className="defaultContainer">
						{this.getAppDom()}
					</div>
				</Provider>
			);
		}
	}

	public getAppDom = () => {
		return (
			<div className="appContainer">
				<AppRoute />
			</div>);
	}

	/**
	* Renders the component
	*/
	public render(): JSX.Element {
		return (
			<div>
				{this.setThemeComponent()}
			</div>
		);
	}
}