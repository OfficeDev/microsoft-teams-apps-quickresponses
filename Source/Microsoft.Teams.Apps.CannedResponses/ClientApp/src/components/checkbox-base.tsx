// <copyright file="checkbox-base.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Checkbox } from "@fluentui/react";

interface ICheckboxState {
	isCheckboxChecked: boolean;
}
interface ICheckboxProps {
	value: string;
	onCheckboxChecked: (responseId: string, isChecked: boolean) => void;
}

export default class CheckboxBase extends React.Component<ICheckboxProps, ICheckboxState> {

	constructor(props: ICheckboxProps) {
		super(props);

		this.state = {
			isCheckboxChecked: false
		}
	}

	/**
	*Triggers when user checks/unchecks checkbox to set state
	*/
	onChange = (responseId: string, isChecked: boolean) => {
		this.setState({ isCheckboxChecked: isChecked });
		this.props.onCheckboxChecked(responseId, isChecked);
	}

	/**
	* Renders the component
	*/
	public render(): JSX.Element {
		return (
			<div>
				<Checkbox checked={this.state.isCheckboxChecked} onChange={() => this.onChange(this.props.value, !this.state.isCheckboxChecked)} />
			</div>)
	}
}