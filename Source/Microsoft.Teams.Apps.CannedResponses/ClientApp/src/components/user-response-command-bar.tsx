// <copyright file="command-bar.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Button, Input, Dialog, Flex } from "@fluentui/react";

interface ICommandBarProps {
	isEditEnable: boolean;
	isDeleteEnable: boolean;
	onAddButtonClick: () => void;
	onEditButtonClick: () => void;
	onDeleteButtonClick: () => void;
	handleTableFilter: (searchText: string) => void;
	resoureStrings: any;
}

interface ICommandbarState {
	searchValue: string
}

export default class CommandBar extends React.Component<ICommandBarProps, ICommandbarState> {

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
				<Button icon="add" content={this.props.resoureStrings.addNewButtonText} text className="add-new-button" onClick={() => this.props.onAddButtonClick()} />
				<Button icon="edit" content={this.props.resoureStrings.editButtonText} text disabled={!this.props.isEditEnable} className="edit-button" onClick={() => this.props.onEditButtonClick()} />
				<Dialog
					cancelButton={this.props.resoureStrings.cancelButtonText}
					confirmButton={this.props.resoureStrings.confirmButtonText}
					content={this.props.resoureStrings.dialogConfirmText}
					header={this.props.resoureStrings.dialogConfirmHeader}
					trigger={<Button icon="trash-can" content={this.props.resoureStrings.deleteButtonText} text disabled={!this.props.isDeleteEnable} className="delete-button" />}
					onConfirm={this.props.onDeleteButtonClick}
				/>
				<Flex.Item push>
					<div style={{ width: "40rem" }}>
					<Input
						icon="search"
						fluid placeholder={this.props.resoureStrings.searchPlaceholder}
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