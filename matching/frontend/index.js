import {initializeBlock} from '@airtable/blocks/ui';
import React from 'react';

class CaseCountSummary {
    getData() {
        const base = useBase();
        const table = base.getTableByName('TABLE_NAME');
        const attorneyField = table.getFieldByName('Attorney');
        const statusField = table.getFieldByName('Status');
        const records = useRecords(table, {fields: [attorneyField, statusField]});
        let caseCounts = new Map();
        for (const record of records) {
            const attorneyName = record.getCellValue(attorneyField);
            const attorneyString = attorneyName === null ? 'Unassigned' : record.getCellValueAsString(attorneyField);
            let statusCounts;
            if (!caseCounts.has(attorneyString)) {
                statusCounts = new Map();
                caseCounts.set(attorneyString, statusCounts);
            } else {
                statusCounts = recordsByXValueString.get(attorneyString);
            }
            const statusName = record.getCellValue(statusField);
            const statusString = statusName === null ? 'No Status' : record.getCellValueAsString(statusField);
            let statusCount;
            if (!statusCounts.has(statusString)) {
                statusCount = 0;
            } else {
                statusCount = statusCounts.get(statusString);
            }
            statusCounts.set(statusString, statusCount + 1);
        }
        return caseCounts;
    }
    getHtml() {
        const statusesByTime = [
            'Initial Submission',
            'LL Jotform Submission',
            'Assigned to Attorney',
            'Settled',
            'Done',
            'No Eviction Case Found',
            'No Status'
        ];
        let html = '<table><th><td></td><td>';
        for (const title of statusesByTime) {
            html += (title + '</td>');
        }
        let caseCounts = this.getData();
        for (const attorneyCounts in caseCounts) {
            html += '<tr><td>';
            html += (attorneyCounts.getKey() + '</td><td>';
            for (const count in attorneyCounts) {
                html += (count + '</td><td>')
            }
            html += '</td></tr>';
        }
        html += '</td></th>';
        html += '</table>';
        return html;
    }
}

function HelloWorldApp() {
    return (new CaseCountSummary().getHtml());
}

initializeBlock(() => <HelloWorldApp />);
