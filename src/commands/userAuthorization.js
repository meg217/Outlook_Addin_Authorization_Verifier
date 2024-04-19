//import * as fs from 'node:fs/promises';
//const csv = require('csv-parser');

// This function checks if the user's clearance meets requirements
function userMeetsSecurityClearance(filePath, documentClassification, email) {
    return new Promise((resolve, reject) => {
        let accessGranted = false;
        let email = email.toLowerCase();
        console.log("userMeetsSecurityClearance Function, checking for email: ", email)
        // Fetch the CSV file
        fetch(filePath)
        .then(response => response.text())
        .then(csvData => {
            Papa.parse(csvData, {
                header: true,
                complete: (results) => {
                    results.data.forEach(row => {
                        console.log("Email: ", row["Email"]);
                        if (row["Email"] === email) {
                            console.log("Found email in row: ", row);
                            const userClearance = row["Authorization"];

                            if (canUserAccess(documentClassification, userClearance)) {
                                accessGranted = true;
                                console.log("accessGranted = true");
                            }
                        }
                    });
                    resolve(accessGranted);
                },
                error: (error) => {
                    console.error("Error parsing CSV:", error);
                    reject(error);
                }
            });
        })
        .catch(error => {
            console.error("Error fetching CSV:", error);
            reject(error);
        });
    });
}
function canUserAccess(documentClassification, userClearance) {
    console.log("canUserAccess Function")
    const levels = ['confidential', 'secret', 'top secret'];
    const documentIndex = levels.indexOf(documentClassification.trim().toLowerCase());
    const userIndex = levels.indexOf(userClearance.trim().toLowerCase());

    return userIndex >= documentIndex;
}

// Example usage
//const filePath = 'users.csv'; // Adjust as needed
//const documentClassification = 'secret';
//const email = 'johndoe@yahoo.tx.gov';

//userMeetsSecurityClearance(filePath, documentClassification, email)
//    .then((result) => {
//        console.log(result); // true or false
//    })
//    .catch((error) => {
//        console.error(`An error occurred: ${error}`);
//    });
