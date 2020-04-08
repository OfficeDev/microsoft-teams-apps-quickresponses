// <copyright file="company-responses-command-bar.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Button, Input, Flex } from "@fluentui/react";

interface ICommandBarProps {
	showSuggestNewButton: boolean;
	onSuggestNewButtonClick: () => void;
	handleTableFilter: (searchText: string) => void;
	resourceStrings:any
}

interface ICommandbarState {
	searchValue: string
}

export default class CompanyResponsesCommandBar extends React.Component<ICommandBarProps, ICommandbarState> {

	constructor(props: ICommandBarProps) {
		super(props);
		this.state = { searchValue: "" };
		this.handleChange = this.handleChange.bind(this);
		this.handleKeyPress = this.handleKeyPress.bind(this);
	}

	/**
	* Set State value of text box input control
	* @param  {Any} event Event object
	*/
	handleChange(event: any) {
		this.setState({ searchValue: event.target.value });
		if (event.target.value.length > 2 || event.target.value === "") {
			this.props.handleTableFilter(event.target.value);
		}
	}

	/**
	* Used to call parent search method on enter key press in text box
	* @param  {Any} event Event object
	*/
	handleKeyPress(event: any) {
		var keyCode = event.which || event.keyCode;
		if (keyCode == 13) {
			if (event.target.value.length > 2 || event.target.value === "") {
				this.props.handleTableFilter(event.target.value);
			}
		}
	}

	/**
	* Renders the component
	*/
	public render(): JSX.Element {
		return (
			<Flex gap="gap.small" className="commandbar-wrapper">
				{this.props.showSuggestNewButton === true && <Button icon="add" content={this.props.resourceStrings.suggestNewButtonText} text className="add-new-button" onClick={() => this.props.onSuggestNewButtonClick()} />}
				<Flex.Item push>
					<div style={{ width: "40rem" }}>
					<Input
						icon="search"
						fluid
						placeholder={this.props.resourceStrings.searchPlaceholder}
						value={this.state.searchValue}
						onChange={this.handleChange}
						onKeyUp={this.handleKeyPress}
						/>
					</div>
				</Flex.Item>
			</Flex>
		);
	}
}