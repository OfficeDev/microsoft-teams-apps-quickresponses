// <copyright file="user-requests-table.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Table, Text } from "@fluentui/react";
var moment = require('moment');

interface IUserRequestsTableProps {
    requestsData: any[],
    resoureStrings: any
}

const UserRequestsTable: React.FunctionComponent<IUserRequestsTableProps> = props => {
    const userRequestsTableHeader = {
        key: "header",
        items: [
            { content: <Text weight="regular" content={props.resoureStrings.response} />, key: "response" },
            { content: <Text weight="regular" content={props.resoureStrings.questions} />, key: "questions" },
            { content: <Text weight="regular" content={props.resoureStrings.label} />, key: "label", className: "table-label-cell" },
            { content: <Text weight="regular" content={props.resoureStrings.requestedOnText} />, key: "requestedon", className: "table-label-cell" },
            { content: <Text weight="regular" content={props.resoureStrings.statusText} />, key: "status", className: "table-label-cell" }
        ],
    };

    let userRequestsTableRows = props.requestsData.map((value: any, index) => (
        {
            key: index,
            items: [
                { content: <Text content={value.responseText} title={value.responseText} />, key: index + "2", truncateContent: true },
                { content: <Text content={value.questionText} title={value.questionText} />, key: index + "3", truncateContent: true },
                { content: <Text content={value.questionLabel} title={value.questionLabel} />, key: index + "4", truncateContent: true, className: "table-label-cell" },
                { content: <Text content={moment.utc(value.createdDate).local().format("MM-DD-YYYY hh:mm A")} title={moment.utc(value.createdDate).local().format("MM-DD-YYYY hh:mm A")} />, key: index + "5", truncateContent: true, className: "table-label-cell" },
                {
                    content: <Text content={value.approvalStatus} styles={{ color: value.approvalStatus === "Approved" ? "#237B4B" : value.approvalStatus === "Pending" ? "#FFAA44" : "#8E192E" }} title={value.approvalStatus} />,
                    key: index + "6", truncateContent: true, className: "table-label-cell"
                },
            ],
        }
    ));

    return (
        <div>
            <Table rows={userRequestsTableRows}
                header={userRequestsTableHeader} className="table-cell-content" />
        </div>
    );
}

export default UserRequestsTable;