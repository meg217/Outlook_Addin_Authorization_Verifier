//import * as fs from 'node:fs/promises';
//const csv = require('csv-parser');

// This function checks if the user's clearance meets requirements
async function userMeetsSecurityClearance(filePath, documentClassification, email1) {
    let accessGranted = false;
    let email = email1.toLowerCase();
    console.log("userMeetsSecurityClearance Function, checking for email: ", email);

    try {
        const response = await fetch(filePath);
        const csvData = await response.text();
        const results = Papa.parse(csvData, { header: true }).data;

        let foundEmail = false;
        for (const row of results) {
            if (row["Email"] === email) {
                console.log("Found email in row: ", row);
                foundEmail = true;
                const userClearance = row["Authorization"];
                if (canUserAccess(documentClassification, userClearance)) {
                    accessGranted = true;
                    console.log("accessGranted = true");
                }
            }
        }

        if (!foundEmail) {
            throw new Error("Email not found in CSV");
        }

        return accessGranted;
    } catch (error) {
        console.error("Error:", error);
        throw error;
    }
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
