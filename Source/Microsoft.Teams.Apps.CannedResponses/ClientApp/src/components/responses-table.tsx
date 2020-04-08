// <copyright file="responses-table.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Table, Text } from "@fluentui/react";
import CheckboxBase from "./checkbox-base";
import "../styles/site.css";

interface IResponsesTableProps {
    showCheckbox: boolean,
    responsesData: any[],
    onCheckBoxChecked: (responseId: string, isChecked: boolean) => void,
    resoureStrings:any
}

const ResponsesTable: React.FunctionComponent<IResponsesTableProps> = props => {
    const userResponsesTableHeader = {
        key: "header",
        items: props.showCheckbox === true ?
            [
                { content: <div />, key: "check-box", className:"table-checkbox-cell" },
                {
                    content: <Text weight="regular" content={props.resoureStrings.response} />, key: "response"
                },
                { content: <Text weight="regular" content={props.resoureStrings.questions} />, key: "questions" },
                { content: <Text weight="regular" content={props.resoureStrings.label} />, key: "label", className: "table-label-cell" }
            ]
            :
            [
                { content: <Text weight="regular" content={props.resoureStrings.response} />, key: "response" },
                { content: <Text weight="regular" content={props.resoureStrings.questions} />, key: "questions" },
                { content: <Text weight="regular" content={props.resoureStrings.label} />, key: "label", className:"table-label-cell" }
            ],
    };

    let UserResponsesTableRows = props.responsesData.map((value: any, index) => (
        {
            key: index,
            style: {},
            items: props.showCheckbox === true ?
                [
                    { content: <CheckboxBase onCheckboxChecked={props.onCheckBoxChecked} value={value.responseId} />, key: index + "1", className: "table-checkbox-cell"},
                    { content: <Text content={value.responseText} title={value.responseText} />, key: index + "2", truncateContent: true },
                    { content: <Text content={value.questionText} title={value.questionText} />, key: index + "3", truncateContent: true },
                    { content: <Text content={value.questionLabel} title={value.questionLabel} />, key: index + "4", truncateContent: true, className: "table-label-cell" },
                ]
                :
                [
                    { content: <Text content={value.responseText} title={value.responseText} />, key: index + "2", truncateContent: true },
                    { content: <Text content={value.questionText} title={value.questionText} />, key: index + "3", truncateContent: true },
                    { content: <Text content={value.questionLabel} title={value.questionLabel} />, key: index + "4", truncateContent: true, className: "table-label-cell" },
                ],
        }
    ));

    return (
        <div>
            <Table rows={UserResponsesTableRows}
                header={userResponsesTableHeader} className="table-cell-content" />
        </div>
    );
}

export default ResponsesTable;