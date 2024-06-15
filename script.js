const baseId = 'appexR9tFKGHjSWNE';
const apiKey = 'patvDa3N84MeSmPqS.bea18731fee485325ba61b363719bf9188151c9932245d8f43e7cc56c645fdf2';

function formatDate(dateString) {
    const date = new Date(dateString);
    const options = {
        year: 'numeric',
        month: 'numeric',
        day: 'numeric',
        hour: 'numeric',
        minute: 'numeric',
        hour12: true
    };
    return date.toLocaleString('en-US', options);
}

function updateAirtableInfoTable() {
    const apiUrl = `https://api.airtable.com/v0/${baseId}/${encodeURIComponent('Unified Desklogs')}`;
    fetch(apiUrl, {
        headers: {
            'Authorization': 'Bearer ' + apiKey
        }
    })
    .then(response => response.json())
    .then(data => {
        const tableData = data.records.map(record => ({
            dealership: record.fields['Dealership'] || 'N/A',
            confirmed: record.fields['Confirmed'] || 0,
            shown: record.fields['Shown'] || 0,
            missed: record.fields['Missed'] || 0,
            overdue: record.fields['Overdue'] || 0,
            internet: record.fields['Internet'] || 0,
            remainingToday: record.fields['Remaining Today'] || 0,
            the1st48Today: record.fields['The 1st48 Today'] || 0,
            shownPercentage: parseFloat(record.fields['Shown Percentage']) || 0,
            today: record.fields['Appointments for Today'] || 0,
            tomorrow: record.fields['Appointments for Tomorrow'] || 0,
            thirdDay: record.fields['3rd Day Appointments'] || 0
        }));

        $('#airtable-info-table').DataTable({
            data: tableData,
            columns: [
                { title: "Dealership", data: "dealership" },
                { title: "Confirmed", data: "confirmed" },
                { title: "Shown", data: "shown" },
                { title: "Missed", data: "missed" },
                { title: "Overdue", data: "overdue" },
                { title: "Internet", data: "internet" },
                { title: "Today", data: "today" },
                { title: "Tomorrow", data: "tomorrow" },
                { title: "3rd Day", data: "thirdDay" },
                { title: "Remaining Today", data: "remainingToday" },
                { title: "The 1st48 Today", data: "the1st48Today" },
                { title: "Shown Percentage", data: "shownPercentage" }
            ],
            pageLength: 20,
            order: [[11, 'desc']],
            destroy: true,
            initComplete: function() {
                dataTablesReady.airtable = true;
                checkDataTablesReady();
            }
        });
    })
    .catch(error => console.error('Error loading the data:', error));
}

function updateAppointmentsInfoTable() {
    const apiUrl = `https://api.airtable.com/v0/${baseId}/${encodeURIComponent('Appointments')}`;
    fetch(apiUrl, {
        headers: {
            'Authorization': 'Bearer ' + apiKey
        }
    })
    .then(response => response.json())
    .then(data => {
        const tableData = data.records.map(record => {
            const fields = record.fields;
            return {
                customerName: fields['Customer Name'] || 'N/A',
                customerPhone: fields['Customer Phone'] || 'N/A',
                appointmentDate: fields['Appointment Date'] ? formatDate(fields['Appointment Date']) : 'N/A',
                apptOrDFU: fields['Appt or DFU'] || 'N/A',
                comment: fields['Comment'] || 'N/A',
                created: new Date(record.createdTime)
            };
        });

        $('#appointments-info-table').DataTable({
            data: tableData,
            columns: [
                { title: "Customer Name", data: "customerName" },
                { title: "Customer Phone", data: "customerPhone" },
                { title: "Appointment Date", data: "appointmentDate" },
                { title: "Appt or DFU", data: "apptOrDFU" },
                { title: "Comment", data: "comment" },
                { title: "Created", data: "created" }
            ],
            pageLength: 20,
            order: [[5, 'desc']],
            destroy: true,
            initComplete: function() {
                countAppointmentsAndFollowUps();
            }
        });
    })
    .catch(error => console.error('Error loading Appointments data:', error));
}

function updateVinSoInfoTable() {
    const apiUrl = `https://api.airtable.com/v0/${baseId}/${encodeURIComponent('VinSo Desklog')}`;
    fetch(apiUrl, {
        headers: {
            'Authorization': 'Bearer ' + apiKey
        }
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        console.log('VinSo data received:', data); // Log the data to check the response
        const tableData = data.records.map(record => {
            const fields = record.fields;
            return {
                name: fields && fields['Name'] ? fields['Name'] : 'N/A',
                inbounds: fields && fields['Inbounds'] ? Number(fields['Inbounds']) : 0,
                outbounds: fields && fields['Outbounds'] ? Number(fields['Outbounds']) : 0,
                texts: fields && fields['Texts'] ? Number(fields['Texts']) : 0,
                emails: fields && fields['Emails'] ? Number(fields['Emails']) : 0,
                dueTasks: fields && fields['DueTasks'] ? Number(fields['DueTasks']) : 0,
                overDue: fields && fields['OverDue'] ? Number(fields['OverDue']) : 0,
                sold: fields && fields['Sold'] ? Number(fields['Sold']) : 0,
                updated: fields && fields['Updated'] ? formatDate(fields['Updated']) : 'N/A'
            };
        });

        console.log('VinSo table data:', tableData); // Log the processed table data
        $('#vinso-info-table').DataTable({
            data: tableData,
            columns: [
                { title: "Name", data: "name" },
                { title: "Inbounds", data: "inbounds" },
                { title: "Outbounds", data: "outbounds" },
                { title: "Texts", data: "texts" },
                { title: "Emails", data: "emails" },
                { title: "Due Tasks", data: "dueTasks" },
                { title: "Overdue", data: "overDue" },
                { title: "Sold", data: "sold" },
                { title: "Updated", data: "updated" }
            ],
            destroy: true,
            initComplete: function() {
                dataTablesReady.vinso = true;
                checkDataTablesReady();
            }
        });
    })
    .catch(error => console.error('Error loading VinSo data:', error));
}

let appointmentsCount = 0;
let followUpsCount = 0;
let dataTablesReady = {
    airtable: false,
    vinso: false,
    appointments: false
};

function countAppointmentsAndFollowUps() {
    const table = $('#appointments-info-table').DataTable();
    const tableData = table.rows().data().toArray();

    appointmentsCount = tableData.filter(item => item.apptOrDFU === 'Appointment').length;
    followUpsCount = tableData.filter(item => item.apptOrDFU === 'Follow-up').length;

    console.log('Appointments count:', appointmentsCount);
    console.log('Follow-ups count:', followUpsCount);

    dataTablesReady.appointments = true;
    checkDataTablesReady();
}

function checkDataTablesReady() {
    if (dataTablesReady.airtable && dataTablesReady.vinso && dataTablesReady.appointments) {
        createHighchartsChartFromTable();
    }
}

function createHighchartsChartFromTable() {
    const vinsoTable = $('#vinso-info-table').DataTable();
    const vinsoData = vinsoTable.rows().data().toArray();

    console.log('DataTable data for Highcharts:', vinsoData); // Log the data from DataTable

    if (!vinsoData.length) {
        console.error('No data available to create Highcharts chart.');
        return;
    }

    // Create a unified array for all categories and values
    const categories = vinsoData.map(item => item.name);
    const inbounds = vinsoData.map(item => item.inbounds);
    const outbounds = vinsoData.map(item => item.outbounds);
    const texts = vinsoData.map(item => item.texts);
    const emails = vinsoData.map(item => item.emails);
    const dueTasks = vinsoData.map(item => item.dueTasks);
    const overDue = vinsoData.map(item => item.overDue);
    const sold = vinsoData.map(item => item.sold);

    Highcharts.chart('dealership-performance-chart', {
        chart: {
            type: 'column',
            animation: Highcharts.svg
        },
        title: {
            text: 'The 1st48 Performance'
        },
        xAxis: {
            categories: categories,
            crosshair: true
        },
        yAxis: {
            min: 0,
            title: {
                text: 'Values'
            }
        },
        tooltip: {
            headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
            pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                '<td style="padding:0"><b>{point.y:.1f}</b></td></tr>',
            footerFormat: '</table>',
            shared: true,
            useHTML: true
        },
        plotOptions: {
            column: {
                pointPadding: 0.2,
                borderWidth: 0,
                dataLabels: {
                    enabled: true,
                    format: '{point.y:.1f}'
                }
            }
        },
        series: [{
            name: 'Appointments',
            data: new Array(categories.length).fill(appointmentsCount)
        }, {
            name: 'Follow-up',
            data: new Array(categories.length).fill(followUpsCount)
        }, {
            name: 'Inbounds',
            data: inbounds
        }, {
            name: 'Outbounds',
            data: outbounds
        }, {
            name: 'Texts',
            data: texts
        }, {
            name: 'Emails',
            data: emails
        }, {
            name: 'Due Tasks',
            data: dueTasks
        }, {
            name: 'Overdue',
            data: overDue
        }, {
            name: 'Sold',
            data: sold
        }]
    });
}

function downloadXLSXReport() {
    const workbook = XLSX.utils.book_new();

    // Adding Unified Desklogs
    const airtableData = $('#airtable-info-table').DataTable().rows().data().toArray();
    const airtableSheet = XLSX.utils.json_to_sheet(airtableData);
    XLSX.utils.book_append_sheet(workbook, airtableSheet, "Unified Desklogs");

    // Adding Appointments
    const appointmentsData = $('#appointments-info-table').DataTable().rows().data().toArray();
    const appointmentsSheet = XLSX.utils.json_to_sheet(appointmentsData);
    XLSX.utils.book_append_sheet(workbook, appointmentsSheet, "Appointments");

    // Adding VinSo Desklog
    const vinsoData = $('#vinso-info-table').DataTable().rows().data().toArray();
    const vinsoSheet = XLSX.utils.json_to_sheet(vinsoData);
    XLSX.utils.book_append_sheet(workbook, vinsoSheet, "VinSo Desklog");

    // Writing the workbook
    XLSX.writeFile(workbook, 'Desklog_Report.xlsx');
}

window.addEventListener('load', () => {
    updateAirtableInfoTable();
    updateAppointmentsInfoTable();
    updateVinSoInfoTable(); // Fetch data and create chart with VinSo data
});

document.getElementById('download-xlsx').addEventListener('click', function() {
    downloadXLSXReport();
});
