// Global variables
let file1Data = null;
let file2Data = null;
let file1Sheets = [];
let file2Sheets = [];
let comparisonResult = null;

// DOM Elements
const file1Input = document.getElementById('file1');
const file2Input = document.getElementById('file2');
const sheet1Select = document.getElementById('sheet1');
const sheet2Select = document.getElementById('sheet2');
const plantColumnSelect = document.getElementById('plantColumn');
const materialColumnSelect = document.getElementById('materialColumn');
const compareBtn = document.getElementById('compareBtn');
const mainContent = document.getElementById('mainContent');
const thresholdRadios = document.querySelectorAll('input[name="threshold"]');
const customThresholdGroup = document.getElementById('customThresholdGroup');
const customThresholdInput = document.getElementById('customThreshold');

// Event Listeners
file1Input.addEventListener('change', handleFile1Upload);
file2Input.addEventListener('change', handleFile2Upload);
compareBtn.addEventListener('click', handleComparison);
thresholdRadios.forEach(radio => {
    radio.addEventListener('change', handleThresholdChange);
});

function handleThresholdChange() {
    const selectedValue = document.querySelector('input[name="threshold"]:checked').value;
    customThresholdGroup.style.display = selectedValue === 'custom' ? 'block' : 'none';
}

function getThreshold() {
    const selectedValue = document.querySelector('input[name="threshold"]:checked').value;
    if (selectedValue === 'custom') {
        return parseFloat(customThresholdInput.value);
    }
    return parseFloat(selectedValue);
}

function handleFile1Upload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const data = event.target.result;
            const workbook = XLSX.read(data, { type: 'array' });
            file1Sheets = workbook.SheetNames;
            file1Data = workbook;

            // Update sheet selector
            sheet1Select.innerHTML = '';
            file1Sheets.forEach(sheet => {
                const option = document.createElement('option');
                option.value = sheet;
                option.textContent = sheet;
                sheet1Select.appendChild(option);
            });
            sheet1Select.disabled = false;

            // Load first sheet to populate plant/material column options
            updateColumnSelectors();
            checkIfReadyToCompare();
        } catch (error) {
            showAlert('Error loading File 1: ' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function handleFile2Upload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const data = event.target.result;
            const workbook = XLSX.read(data, { type: 'array' });
            file2Sheets = workbook.SheetNames;
            file2Data = workbook;

            // Update sheet selector
            sheet2Select.innerHTML = '';
            file2Sheets.forEach(sheet => {
                const option = document.createElement('option');
                option.value = sheet;
                option.textContent = sheet;
                sheet2Select.appendChild(option);
            });
            sheet2Select.disabled = false;

            checkIfReadyToCompare();
        } catch (error) {
            showAlert('Error loading File 2: ' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function updateColumnSelectors() {
    if (!file1Data || !sheet1Select.value) return;

    try {
        const sheet = file1Data.Sheets[sheet1Select.value];
        const df = XLSX.utils.sheet_to_json(sheet);

        if (df.length === 0) return;

        const columns = Object.keys(df[0]);

        // plant
        plantColumnSelect.innerHTML = '<option value="">None (No grouping)</option>';
        // material
        materialColumnSelect.innerHTML = '<option value="">None</option>';

        columns.forEach(col => {
            if (col !== 'Key') {
                const opt1 = document.createElement('option');
                opt1.value = col;
                opt1.textContent = col;
                plantColumnSelect.appendChild(opt1);

                const opt2 = opt1.cloneNode(true);
                materialColumnSelect.appendChild(opt2);
            }
        });
        plantColumnSelect.disabled = false;
        materialColumnSelect.disabled = false;
    } catch (error) {
        console.error('Error updating column selectors:', error);
    }
}

sheet1Select.addEventListener('change', updateColumnSelectors);

function checkIfReadyToCompare() {
    const ready = file1Data && file2Data && sheet1Select.value && sheet2Select.value;
    compareBtn.disabled = !ready;
}

function handleComparison() {
    if (!file1Data || !file2Data) {
        showAlert('Please upload both files', 'error');
        return;
    }

    const sheet1 = sheet1Select.value;
    const sheet2 = sheet2Select.value;
    const threshold = getThreshold();
    const plantColumn = plantColumnSelect.value || null;
    const materialColumn = materialColumnSelect.value || null;

    try {
        const df1 = XLSX.utils.sheet_to_json(file1Data.Sheets[sheet1]);
        const df2 = XLSX.utils.sheet_to_json(file2Data.Sheets[sheet2]);

        // Validate Key column
        if (!df1.length || !Object.keys(df1[0]).includes('Key')) {
            showAlert('File 1 must have a "Key" column', 'error');
            return;
        }
        if (!df2.length || !Object.keys(df2[0]).includes('Key')) {
            showAlert('File 2 must have a "Key" column', 'error');
            return;
        }

        // Perform comparison
        comparisonResult = compareDataframes(df1, df2, threshold, plantColumn, materialColumn);

        if (!comparisonResult) {
            showAlert('No matching Keys found between files', 'error');
            return;
        }

        // Display results
        displayResults(comparisonResult, threshold, plantColumn, materialColumn);
        showAlert(`✅ Comparison complete - ${comparisonResult.length} records with matching Keys found`, 'success');
    } catch (error) {
        showAlert('Error during comparison: ' + error.message, 'error');
    }
}

function compareDataframes(df1, df2, threshold, plantColumn, materialColumn) {
    // Create index maps
    const df2Map = {};
    df2.forEach(row => {
        df2Map[row.Key] = row;
    });

    const comparisonData = [];
    const numericCols = getNumericColumns(df1);

    df1.forEach(row1 => {
        const key = row1.Key;
        if (!df2Map[key]) return;

        const row2 = df2Map[key];
        const comparisonRow = { Key: key };

        if (plantColumn && row1[plantColumn]) {
            comparisonRow[plantColumn] = row1[plantColumn];
        }
        if (materialColumn && row1[materialColumn]) {
            comparisonRow[materialColumn] = row1[materialColumn];
        }

        let hasOutlier = false;

        numericCols.forEach(col => {
            if (row1[col] === undefined || row2[col] === undefined) return;

            const val1 = parseFloat(row1[col]);
            const val2 = parseFloat(row2[col]);

            if (isNaN(val1) || isNaN(val2)) return;

            const pctDiff = val1 !== 0 ? Math.abs((val2 - val1) / val1) * 100 : (val2 === 0 ? 0 : 100);

            comparisonRow[`${col}_File1`] = val1;
            comparisonRow[`${col}_File2`] = val2;
            comparisonRow[`${col}_Diff%`] = pctDiff.toFixed(2);
            comparisonRow[`${col}_IsOutlier`] = pctDiff > threshold;

            if (pctDiff > threshold) {
                hasOutlier = true;
            }
        });

        if (Object.keys(comparisonRow).length > 1) {
            comparisonRow.hasOutlier = hasOutlier;
            comparisonData.push(comparisonRow);
        }
    });

    return comparisonData.length > 0 ? comparisonData : null;
}

function getNumericColumns(df) {
    if (!df.length) return [];
    const columns = Object.keys(df[0]);
    const numericCols = [];

    columns.forEach(col => {
        if (col !== 'Key') {
            const isNumeric = df.some(row => {
                const val = parseFloat(row[col]);
                return !isNaN(val);
            });
            if (isNumeric) {
                numericCols.push(col);
            }
        }
    });

    return numericCols;
}

function displayResults(data, threshold, plantColumn, materialColumn) {
    const numericCols = getNumericColumnsFromComparison(data);

    let html = `
        <div class="tabs">
            <button class="tab-button active" onclick="switchTab(event, 'overview')">📊 Overview</button>
            <button class="tab-button" onclick="switchTab(event, 'outliers')">🎯 Outliers</button>
    `;

    if (plantColumn) {
        html += `<button class="tab-button" onclick="switchTab(event, 'plants')">🏭 Plant Summary</button>`;
    }

    html += `
            <button class="tab-button" onclick="switchTab(event, 'detailed')">📋 Detailed View</button>
        </div>

        <!-- Overview Tab -->
        <div id="overview" class="tab-content active">
            <div class="metrics">
                ${createMetricsCards(data, numericCols)}
            </div>
            <div class="chart-container">
                <canvas id="distributionChart"></canvas>
            </div>
        </div>

        <!-- Outliers Tab -->
        <div id="outliers" class="tab-content">
            ${createOutliersContent(data, numericCols, threshold)}
        </div>
    `;

    if (plantColumn) {
        html += `
        <!-- Plant Summary Tab -->
        <div id="plants" class="tab-content">
            ${createPlantSummary(data, plantColumn, materialColumn)}
        </div>
        `;
    }

    html += `
        <!-- Detailed Tab -->
        <div id="detailed" class="tab-content">
            ${createDetailedTab(data, numericCols, plantColumn, materialColumn)}
        </div>
    `;

    mainContent.innerHTML = html;

    // Create charts
    setTimeout(() => {
        createDistributionChart(data, numericCols);
        if (document.getElementById('outliersChart')) {
            createOutliersChart(data, numericCols);
        }
        if (document.getElementById('topOutliersChart')) {
            createTopOutliersChart(data, numericCols);
        }
    }, 100);
}

function createMetricsCards(data, numericCols) {
    let totalOutliers = 0;
    let allDiffs = [];

    data.forEach(row => {
        numericCols.forEach(col => {
            const diffKey = `${col}_Diff%`;
            const outlierKey = `${col}_IsOutlier`;
            if (row[outlierKey]) totalOutliers++;
            if (row[diffKey] !== undefined) {
                allDiffs.push(parseFloat(row[diffKey]));
            }
        });
    });

    const avgDiff = allDiffs.length > 0 ? (allDiffs.reduce((a, b) => a + b) / allDiffs.length).toFixed(2) : 0;
    const maxDiff = allDiffs.length > 0 ? Math.max(...allDiffs).toFixed(2) : 0;

    return `
        <div class="metric-card">
            <h3>Total Records</h3>
            <div class="value">${data.length}</div>
        </div>
        <div class="metric-card">
            <h3>Total Outliers</h3>
            <div class="value">${totalOutliers}</div>
        </div>
        <div class="metric-card">
            <h3>Avg Difference %</h3>
            <div class="value">${avgDiff}%</div>
        </div>
        <div class="metric-card">
            <h3>Max Difference %</h3>
            <div class="value">${maxDiff}%</div>
        </div>
    `;
}

function createOutliersContent(data, numericCols, threshold) {
    const outlierData = data.filter(row => row.hasOutlier);

    if (outlierData.length === 0) {
        return '<div class="no-data">✅ No outliers found above the threshold!</div>';
    }

    let html = `
        <div class="alert alert-info">
            🎯 Found ${outlierData.length} records with outliers (threshold: ${threshold}%)
        </div>
        <div class="metrics">
            <div class="chart-container">
                <canvas id="outliersChart"></canvas>
            </div>
            <div class="chart-container">
                <div id="topOutliersChart" style="height: 400px;"></div>
            </div>
        </div>
        <h3 style="margin-top: 30px;">Outlier Records</h3>
        ${createOutlierTable(outlierData, numericCols)}
    `;

    return html;
}

function createOutlierTable(data, numericCols) {
    let html = '<table><thead><tr><th>Key</th>';

    numericCols.forEach(col => {
        html += `<th>${col} (File1)</th><th>${col} (File2)</th><th>${col} Diff%</th>`;
    });

    html += '</tr></thead><tbody>';

    data.forEach(row => {
        html += `<tr class="outlier">`;
        html += `<td>${row.Key}</td>`;

        numericCols.forEach(col => {
            const file1Val = row[`${col}_File1`]?.toFixed(2) || '-';
            const file2Val = row[`${col}_File2`]?.toFixed(2) || '-';
            const diff = row[`${col}_Diff%`] || '-';
            const isOutlier = row[`${col}_IsOutlier`];

            html += `<td>${file1Val}</td>`;
            html += `<td>${file2Val}</td>`;
            html += `<td class="${isOutlier ? 'outlier-cell' : ''}">${diff}${diff !== '-' ? '%' : ''}</td>`;
        });

        html += '</tr>';
    });

    html += '</tbody></table>';
    return html;
}

function createPlantSummary(data, plantColumn, materialColumn) {
    const plantSummary = {};

    data.forEach(row => {
        const plant = row[plantColumn] || 'Unknown';
        if (!plantSummary[plant]) {
            plantSummary[plant] = {
                records: 0,
                outliers: 0,
                materials: new Set()
            };
        }
        plantSummary[plant].records++;
        if (row.hasOutlier) {
            plantSummary[plant].outliers++;
        }
        if (materialColumn && row[materialColumn]) {
            plantSummary[plant].materials.add(row[materialColumn]);
        }
    });

    let html = '<div class="plant-summary">';

    Object.entries(plantSummary).forEach(([plant, stats]) => {
        const outlierPct = ((stats.outliers / stats.records) * 100).toFixed(1);
        let matsHtml = '';
        if (materialColumn) {
            matsHtml = `<div class="plant-stat">
                            <span>Materials:</span>
                            <span class="plant-stat-value">${stats.materials.size}</span>
                        </div>`;
        }
        html += `
            <div class="plant-card">
                <h3>🏭 ${plant}</h3>
                <div class="plant-stat">
                    <span>Records:</span>
                    <span class="plant-stat-value">${stats.records}</span>
                </div>
                <div class="plant-stat">
                    <span>Outliers:</span>
                    <span class="plant-stat-value">${stats.outliers}</span>
                </div>
                <div class="plant-stat">
                    <span>Outlier %:</span>
                    <span class="plant-stat-value">${outlierPct}%</span>
                </div>
                ${matsHtml}
            </div>
        `;
    });

    html += '</div>';

    // Add charts
    html += '<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">';
    html += '<div class="chart-container"><canvas id="recordsChart"></canvas></div>';
    html += '<div class="chart-container"><canvas id="outliersPercentChart"></canvas></div>';
    html += '</div>';

    setTimeout(() => {
        createPlantCharts(plantSummary);
    }, 100);

    return html;
}

function createDetailedTab(data, numericCols, plantColumn, materialColumn) {
    let html = `
        <div class="filter-section">
            <div class="filter-row">
                <div>
                    <label>
                        <input type="checkbox" id="outlierFilter" /> Show outliers only
                    </label>
                </div>
                ${plantColumn ? `
                <div>
                    <label>Filter by Plant:</label>
                    <select id="plantFilter">
                        <option value="">All Plants</option>
                    </select>
                </div>
                ` : ''}
                ${materialColumn ? `
                <div>
                    <label>Filter by Material:</label>
                    <select id="materialFilter">
                        <option value="">All Materials</option>
                    </select>
                </div>
                ` : ''}
                <div>
                    <button onclick="applyDetailedFilters()" style="background: #667eea;">🔍 Apply Filters</button>
                </div>
            </div>
        </div>
        <div id="detailedTable"></div>
        <button class="download-btn" onclick="downloadCSV()">📥 Download as CSV</button>
    `;

    // Populate plant filter
    if (plantColumn) {
        setTimeout(() => {
            const plants = [...new Set(data.map(r => r[plantColumn]))].filter(p => p);
            const plantFilter = document.getElementById('plantFilter');
            plants.forEach(plant => {
                const option = document.createElement('option');
                option.value = plant;
                option.textContent = plant;
                plantFilter.appendChild(option);
            });
        }, 50);
    }
    if (materialColumn) {
        setTimeout(() => {
            const mats = [...new Set(data.map(r => r[materialColumn]))].filter(m => m);
            const matFilter = document.getElementById('materialFilter');
            mats.forEach(mat => {
                const option = document.createElement('option');
                option.value = mat;
                option.textContent = mat;
                matFilter.appendChild(option);
            });
        }, 50);
    }

    setTimeout(() => {
        displayDetailedTable(data, numericCols, plantColumn, materialColumn);
    }, 100);

    return html;
}

function displayDetailedTable(data, numericCols, plantColumn, materialColumn, filtered = null) {
    const displayData = filtered || data;
    let html = '<table><thead><tr><th>Key</th>';

    if (plantColumn) {
        html += `<th>${plantColumn}</th>`;
    }
    if (materialColumn) {
        html += `<th>${materialColumn}</th>`;
    }

    numericCols.slice(0, 3).forEach(col => {
        html += `<th>${col} (Diff%)</th>`;
    });

    html += '</tr></thead><tbody>';

    displayData.forEach(row => {
        const rowClass = row.hasOutlier ? 'outlier' : '';
        html += `<tr class="${rowClass}">`;
        html += `<td>${row.Key}</td>`;

        if (plantColumn) {
            html += `<td>${row[plantColumn] || '-'}</td>`;
        }

        numericCols.slice(0, 3).forEach(col => {
            const diff = row[`${col}_Diff%`] || '-';
            const isOutlier = row[`${col}_IsOutlier`];
            html += `<td class="${isOutlier ? 'outlier-cell' : ''}">${diff}${diff !== '-' ? '%' : ''}</td>`;
        });

        html += '</tr>';
    });

    html += '</tbody></table>';

    const tableContainer = document.getElementById('detailedTable');
    if (tableContainer) {
        tableContainer.innerHTML = html;
    }
}

function applyDetailedFilters() {
    const outlierOnly = document.getElementById('outlierFilter')?.checked || false;
    const plantFilter = document.getElementById('plantFilter')?.value || '';

    let filtered = comparisonResult;

    if (outlierOnly) {
        filtered = filtered.filter(row => row.hasOutlier);
    }

    if (plantFilter) {
        filtered = filtered.filter(row => row[document.getElementById('plantColumn').value] === plantFilter);
    }
    const materialFilter = document.getElementById('materialFilter')?.value || '';
    if (materialFilter) {
        filtered = filtered.filter(row => row[document.getElementById('materialColumn').value] === materialFilter);
    }

    const numericCols = getNumericColumnsFromComparison(comparisonResult);
    displayDetailedTable(comparisonResult, numericCols, document.getElementById('plantColumn').value, document.getElementById('materialColumn').value, filtered);
}

function switchTab(e, tabName) {
    // Hide all tabs
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });

    // Remove active class from all buttons
    document.querySelectorAll('.tab-button').forEach(btn => {
        btn.classList.remove('active');
    });

    // Show selected tab
    const tab = document.getElementById(tabName);
    if (tab) {
        tab.classList.add('active');
    }

    // Add active class to clicked button
    e.target.classList.add('active');
}

function getNumericColumnsFromComparison(data) {
    if (!data || data.length === 0) return [];

    const columns = Object.keys(data[0]);
    return [...new Set(
        columns
            .filter(col => col.endsWith('_File1'))
            .map(col => col.replace('_File1', ''))
    )];
}

function createDistributionChart(data, numericCols) {
    const ctx = document.getElementById('distributionChart');
    if (!ctx) return;

    const allDiffs = [];
    data.forEach(row => {
        numericCols.forEach(col => {
            const diff = parseFloat(row[`${col}_Diff%`]);
            if (!isNaN(diff)) {
                allDiffs.push(diff);
            }
        });
    });

    // Create histogram bins
    const bins = [];
    const binSize = 5;
    for (let i = 0; i <= 100; i += binSize) {
        bins.push(0);
    }

    allDiffs.forEach(diff => {
        const binIndex = Math.floor(diff / binSize);
        if (binIndex < bins.length) {
            bins[binIndex]++;
        }
    });

    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: bins.map((_, i) => `${i * binSize}-${(i + 1) * binSize}%`),
            datasets: [{
                label: 'Difference Distribution',
                data: bins,
                backgroundColor: 'rgba(102, 126, 234, 0.7)',
                borderColor: 'rgba(102, 126, 234, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Distribution of Percentage Differences'
                }
            },
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

function createOutliersChart(data, numericCols) {
    const ctx = document.getElementById('outliersChart');
    if (!ctx) return;

    const outlierCounts = {};
    numericCols.forEach(col => {
        outlierCounts[col] = 0;
    });

    data.forEach(row => {
        if (row.hasOutlier) {
            numericCols.forEach(col => {
                if (row[`${col}_IsOutlier`]) {
                    outlierCounts[col]++;
                }
            });
        }
    });

    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: Object.keys(outlierCounts),
            datasets: [{
                label: 'Outlier Count by Metric',
                data: Object.values(outlierCounts),
                backgroundColor: 'rgba(245, 87, 108, 0.7)',
                borderColor: 'rgba(245, 87, 108, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Outliers by Metric'
                }
            },
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

function createTopOutliersChart(data, numericCols) {
    const chartDiv = document.getElementById('topOutliersChart');
    if (!chartDiv) return;

    const topOutliers = [];

    data.forEach(row => {
        numericCols.forEach(col => {
            if (row[`${col}_IsOutlier`]) {
                topOutliers.push({
                    key: row.Key,
                    metric: col,
                    diff: parseFloat(row[`${col}_Diff%`])
                });
            }
        });
    });

    topOutliers.sort((a, b) => b.diff - a.diff);
    const top10 = topOutliers.slice(0, 10);

    const xLabels = top10.map(o => `${o.key} - ${o.metric}`);
    const yValues = top10.map(o => o.diff);

    const trace = {
        x: xLabels,
        y: yValues,
        type: 'bar',
        marker: { color: 'rgba(245, 87, 108, 0.7)' }
    };

    const layout = {
        title: 'Top 10 Outlier Differences',
        xaxis: { title: 'Key - Metric' },
        yaxis: { title: 'Difference %' },
        height: 400
    };

    Plotly.newPlot(chartDiv, [trace], layout, { responsive: true });
}

function createPlantCharts(plantSummary) {
    const plants = Object.keys(plantSummary);
    const records = plants.map(p => plantSummary[p].records);
    const outliers = plants.map(p => plantSummary[p].outliers);
    const outlierPct = plants.map(p => ((plantSummary[p].outliers / plantSummary[p].records) * 100).toFixed(1));

    // Records Chart
    const ctx1 = document.getElementById('recordsChart');
    if (ctx1) {
        new Chart(ctx1, {
            type: 'bar',
            data: {
                labels: plants,
                datasets: [{
                    label: 'Records',
                    data: records,
                    backgroundColor: 'rgba(102, 126, 234, 0.7)',
                    borderColor: 'rgba(102, 126, 234, 1)',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: 'Records by Plant'
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true
                    }
                }
            }
        });
    }

    // Outlier Percentage Chart
    const ctx2 = document.getElementById('outliersPercentChart');
    if (ctx2) {
        new Chart(ctx2, {
            type: 'bar',
            data: {
                labels: plants,
                datasets: [{
                    label: 'Outlier %',
                    data: outlierPct,
                    backgroundColor: 'rgba(245, 87, 108, 0.7)',
                    borderColor: 'rgba(245, 87, 108, 1)',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: 'Outlier % by Plant'
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true
                    }
                }
            }
        });
    }
}

function downloadCSV() {
    if (!comparisonResult) return;

    const numericCols = getNumericColumnsFromComparison(comparisonResult);
    let csv = 'Key';

    const plantCol = document.getElementById('plantColumn').value;
    if (plantCol) {
        csv += ',' + plantCol;
    }

    numericCols.forEach(col => {
        csv += `,${col}_File1,${col}_File2,${col}_Diff%`;
    });

    csv += '\n';

    comparisonResult.forEach(row => {
        csv += row.Key;

        if (plantCol) {
            csv += ',' + (row[plantCol] || '');
        }

        numericCols.forEach(col => {
            const file1 = row[`${col}_File1`] || '';
            const file2 = row[`${col}_File2`] || '';
            const diff = row[`${col}_Diff%`] || '';
            csv += `,${file1},${file2},${diff}`;
        });

        csv += '\n';
    });

    const blob = new Blob([csv], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'excel_comparison.csv';
    a.click();
    window.URL.revokeObjectURL(url);
}

function showAlert(message, type) {
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type}`;
    alertDiv.style.position = 'fixed';
    alertDiv.style.top = '20px';
    alertDiv.style.right = '20px';
    alertDiv.style.zIndex = '1000';
    alertDiv.style.maxWidth = '400px';
    alertDiv.textContent = message;

    document.body.appendChild(alertDiv);

    setTimeout(() => {
        alertDiv.remove();
    }, 4000);
}
