function generateReport() {
    const location = document.getElementById('selLocation').value;
    const month_year = document.getElementById('selMonth').value;
    const organization = document.getElementById('selOrganization').value;

    if (!month_year) {
        alert('Please select a month.');
        return;
    }

    showLoading(true);

    fetch('/generate_report', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ location, month_year, organization })
    })
    .then(r => r.json())
    .then(data => {

        showLoading(false);

        if (!data.success) {
            alert(data.message);
            return;
        }
        renderReport(data.data);
    })
    .catch(err => {
        showLoading(false);
        alert('Request failed: ' + err.message);
    });
}

function renderReport(data) {
    document.getElementById('formSection').style.display = 'none';
    document.getElementById('reportSection').style.display = 'block';

    const container = document.getElementById('reportContainer');

    const days = data.days;
    const dayNames = data.day_names;

    let html = `
    <div class="report-title">Attendance Report</div>
    <table class="attendance-table">
    <thead>
        <tr class="col-header">
            <th class="col-sr">Sr.<br>No.</th>
            <th class="col-name">Name</th>
            <th class="col-desig">Designation</th>
            ${days.map(d => `<th>${d}</th>`).join('')}
            <th>Mandays</th>
            <th>OT</th>
            <th class="col-remarks">Remarks</th>
        </tr>
        <tr class="day-name-row">
            <th></th>
            <th colspan="2" style="text-align:left; padding-left:8px;">
                Location: <strong>${data.location}</strong> &nbsp;|&nbsp;
                ${data.month_label} &nbsp;|&nbsp;
                ${data.organization}
            </th>
            ${dayNames.map(d => `<th>${d.substring(0,3)}</th>`).join('')}
            <th></th>
            <th></th>
            <th></th>
        </tr>
    </thead>
    <tbody>
    `;

    data.rows.forEach(row => {
        const dayCells = row.days.map(val => {
            if (val === 'AB') {
                return `<td class="td-absent">AB</td>`;
            }
            if (val === 'WO') {
                return `<td class="td-wo" style="color:#888;font-style:italic;">WO</td>`;
            }
            
            const num = parseFloat(val);
            const isOT = !isNaN(num) && num > 10.0;
            return `<td class="${isOT ? 'td-overtime' : ''}">${formatDecimalTime(val)}</td>`;
        }).join('');

        html += `
        <tr>
            <td class="td-sr">${row.sr_no}</td>
            <td class="td-name">${row.name}</td>
            <td class="td-desig">${row.designation}</td>
            ${dayCells}
            <td class="td-mandays">${row.mandays}</td>
            <td class="td-ot">${row.ot > 0 ? formatDecimalTime(row.ot) : '0'}</td>
            <td class="td-remarks"></td>
        </tr>`;
    });

    html += `</tbody></table>`;
    container.innerHTML = html;
}

function formatDecimalTime(val) {
    
    if (val === 'AB' || val === null || val === undefined) return 'AB';
    const num = parseFloat(val);
    if (isNaN(num)) return val;
    const h = Math.floor(num);
    const m = Math.round((num - h) * 100);
    return `${String(h).padStart(2,'0')}.${String(m).padStart(2,'0')}`;
}

function goBack() {
    document.getElementById('reportSection').style.display = 'none';
    document.getElementById('formSection').style.display = 'flex';
}

function printReport() {
    window.print();
}

function showLoading(show) {
    document.getElementById('loadingOverlay').style.display =
        show ? 'flex' : 'none';
}

function loadLocations() {

    fetch('/get_locations')
    .then(r => r.json())
    .then(res => {

        let ddl = document.getElementById('selLocation');

        ddl.innerHTML =
            '<option value="">Select Location</option>';

        res.data.forEach(item => {

            ddl.innerHTML +=
            `<option value="${item}">
                ${item}
            </option>`;
        });
    });
}


function loadMonths() {

    const location =
        document.getElementById('selLocation').value;

    fetch(`/get_months?location=${location}`)
    .then(r=>r.json())
    .then(res=>{

        let ddl =
            document.getElementById('selMonth');

        ddl.innerHTML='';

        res.data.forEach(item=>{

            ddl.innerHTML +=
            `<option value="${item.value}">
                ${item.label}
            </option>`;
        });
        loadOrganizations();
    });
}

function downloadExcel() {

    const location =
        document.getElementById('selLocation').value;

    const month_year =
        document.getElementById('selMonth').value;

    const organization =
        document.getElementById('selOrganization').value;

    if (!month_year) {
        alert('Please select month');
        return;
    }

    const url =
        `/download_excel?location=${encodeURIComponent(location)}`
        + `&month_year=${encodeURIComponent(month_year)}`
        + `&organization=${encodeURIComponent(organization)}`;
    window.location.href = url;
}

function loadOrganizations() {

    const location =
        document.getElementById('selLocation').value;

    const month =
        document.getElementById('selMonth').value;

    fetch(
        `/get_organizations?location=${location}&month_year=${month}`
    )
    .then(r=>r.json())
    .then(res=>{
        let ddl =
            document.getElementById('selOrganization');
        ddl.innerHTML='';
        res.data.forEach(item=>{
            ddl.innerHTML +=
            `<option value="${item}">
                ${item}
            </option>`;
        });
    });
}

window.onload = function () {
    loadLocations();
};
