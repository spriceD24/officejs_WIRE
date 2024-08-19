Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Add event listener to the buttons only when the Office context is ready
        document.getElementById("apiCallButton").addEventListener("click", callApi);
        document.getElementById("validateButton").addEventListener("click", validateData);
        document.getElementById("startWizardButton").addEventListener("click", startWizard);
        document.getElementById("backButton").addEventListener("click", goBack);
    }
});

async function callApi() {
    const data = {};
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange("A1:C4").load("values");

            await context.sync();

            for (let i = 0; i < range.values.length; i++) {
                for (let j = 0; j < range.values[i].length; j++) {
                    const cellAddress = String.fromCharCode(65 + j) + (i + 1);
                    data[cellAddress] = range.values[i][j];
                }
            }

            const response = await fetch('http://localhost:5000/upload', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(data)
            });

            if (response.ok) {
                document.getElementById("result").innerText = "Upload successful!";
            } else {
                document.getElementById("result").innerText = "Upload failed!";
            }
        });
    } catch (error) {
        console.error(error);
        document.getElementById("result").innerText = `Error: ${error.message}, data = ${JSON.stringify(data)}`;
    }
}

async function validateData() {
    const data = {};
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange("A1:C4").load("values");

            await context.sync();

            for (let i = 0; i < range.values.length; i++) {
                for (let j = 0; j < range.values[i].length; j++) {
                    const cellAddress = String.fromCharCode(65 + j) + (i + 1);
                    data[cellAddress] = range.values[i][j];
                }
            }

            const response = await fetch('http://localhost:5000/validate', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(data)
            });

            const result = await response.json();

            if (response.ok) {
                document.getElementById("result").innerText = "Validation complete!";
                // Process the validation result and highlight cells
                for (let [cell, status] of Object.entries(result)) {
                    const cellRange = sheet.getRange(cell);
                    if (status === "Error") {
                        cellRange.format.fill.color = "red";
                    } else {
                        cellRange.format.fill.clear(); // Clear any existing highlights
                    }
                }
                await context.sync();
            } else {
                document.getElementById("result").innerText = "Validation failed!";
            }
        });
    } catch (error) {
        console.error(error);
        document.getElementById("result").innerText = `Error: ${error.message}, data = ${JSON.stringify(data)}`;
    }
}

function startWizard() {
    // Hide main content and show wizard content
    document.getElementById("mainContent").style.display = "none";
    document.getElementById("wizardContent").style.display = "block";
}

function goBack() {
    // Hide wizard content and show main content
    document.getElementById("wizardContent").style.display = "none";
    document.getElementById("mainContent").style.display = "block";
}