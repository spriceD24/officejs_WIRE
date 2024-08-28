Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Add event listener to the buttons only when the Office context is ready
        document.getElementById("apiCallButton").addEventListener("click", callApi);
        document.getElementById("validateButton").addEventListener("click", validateData);
        document.getElementById("startWizardButton").addEventListener("click", startWizard);
        document.getElementById("backButton").addEventListener("click", goBack);
        document.getElementById("uploadButton").addEventListener("click", uploadModel);
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


async function uploadModel() {
    try {
        const isOnline = Office.context.platform === Office.PlatformType.OfficeOnline;
        const sliceSize = isOnline ? 4194304 : 2097152; // 4MB for online, 2MB for desktop

        return new Promise((resolve, reject) => {
            Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize }, async (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const myFile = result.value;
                    let sliceCount = myFile.sliceCount;
                    let fileContent = [];

                    // Read all slices
                    for (let i = 0; i < sliceCount; i++) {
                        await new Promise((sliceResolve, sliceReject) => {
                            myFile.getSliceAsync(i, (sliceResult) => {
                                if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                                    fileContent.push(sliceResult.value.data);
                                    sliceResolve();
                                } else {
                                    sliceReject(sliceResult.error);
                                }
                            });
                        });
                    }

                    // Combine all slices into a single Uint8Array
                    const combinedContent = new Uint8Array(fileContent.reduce((acc, val) => [...acc, ...new Uint8Array(val)], []));

                    // Create a Blob from the Uint8Array
                    const blob = new Blob([combinedContent], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

                    // Prepare the FormData object for the POST request
                    const formData = new FormData();
                    formData.append('file', blob, 'workbook.xlsx');

                    try {
                        // Make the HTTP POST request to upload the file
                        const response = await fetch('http://localhost:5000/uploadModel', {
                            method: 'POST',
                            body: formData
                        });

                        if (response.ok) {
                            document.getElementById("result").innerText = "Upload successful!";
                        } else {
                            document.getElementById("result").innerText = "Upload failed!";
                        }
                    } catch (error) {
                        console.error(error);
                        document.getElementById("result").innerText = `Error: ${error.message}`;
                    }

                    // Close the file when we're done with it
                    myFile.closeAsync();
                    resolve();
                } else {
                    reject(new Error(result.error.message));
                }
            });
        });
    } catch (error) {
        console.error(error);
        document.getElementById("result").innerText = `Error: ${error.message}`;
    }
}

// Helper function to log messages with platform information
function logMessage(message) {
    const platform = Office.context.platform === Office.PlatformType.OfficeOnline ? "Excel Online" : "Excel Desktop";
    console.log(`[${platform}] ${message}`);
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
