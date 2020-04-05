// <copyright file="router.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { BrowserRouter, Route, Switch } from "react-router-dom";
import { UserResponsePage } from "../components/user-response";
import { CompanyResponsePage } from "../components/company-responses";
import { AddUserResponse } from "../components/add-user-response";
import { EditUserResponse } from "../components/edit-response";
import ResponseMessage from "../components/response-message";
import { ErrorPage} from "../components/error-page";

export const AppRoute: React.FunctionComponent<{}> = () => {

	return (
		<BrowserRouter>
			<Switch>
				<Route exact path="/user-responses" component={UserResponsePage} />
				<Route exact path="/add-new-response" component={AddUserResponse} />
				<Route exact path="/edit-user-response" component={EditUserResponse} />
				<Route exact path="/company-responses" component={CompanyResponsePage} />
				<Route exact path="/response-message" component={ResponseMessage} />
				<Route exact path="/errorpage" component={ErrorPage} />
				<Route exact path="/errorpage/:id" component={ErrorPage} />
			</Switch>
		</BrowserRouter>

	);
};

