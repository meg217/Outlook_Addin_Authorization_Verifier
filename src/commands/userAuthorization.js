//import * as fs from 'node:fs/promises';
//const csv = require('csv-parser');
//async and promise both dont seem to work, need to maybe make promise in
//with the rest of the promises, function times out with both promise and async
// This function checks if the user's clearance meets requirements
function userMeetsSecurityClearance(filePath, documentClassification, email1) {
    return new Promise((resolve, reject) => {
    let accessGranted = false;
    let email = email1.toLowerCase();
    console.log("userMeetsSecurityClearance Function, checking for email: ", email);

    fetch(filePath)
        .then(response => response.text())
        .then(csvData => {
            const results = Papa.parse(csvData, { header: true }).data;

            let foundEmail = false;
            for (const row of results) {
                if (row["Email"] === email) {
                    console.log("Found email in row: ", row);
                    foundEmail = true;
                    const userClearance = row["Authorization"];
                    if (canUserAccess(documentClassification, userClearance)) {
                        accessGranted = true;
                        console.log("AccessGranted = true");
                        resolve(accessGranted); 
                        return; 
                    }
                }
            }

            if (!foundEmail) {
                console.log("Email not found in CSV");
            }

            resolve(accessGranted); 
        })
        .catch(error => {
            console.error("Error:", error);
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

// function check_NOFORN_Access(filePath, email1) {
//     return new Promise((resolve, reject) => {
//     let accessGranted = false;
//     let email = email1.toLowerCase();
//     console.log("check_NOFORN_Access Function, checking for email: ", email);

//     fetch(filePath)
//         .then(response => response.text())
//         .then(csvData => {
//             const results = Papa.parse(csvData, { header: true }).data;

//             let foundEmail = false;
//             for (const row of results) {
//                 if (row["Email"] === email) {
//                     console.log("Found email in row: ", row);
//                     foundEmail = true;
//                     const userCountry = row["Country"];
//                     if (userCountry == "USA") {
//                         accessGranted = true;
//                         console.log("AccessGranted = true");
//                         resolve(accessGranted); 
//                         return; 
//                     }
//                 }
//             }

//             if (!foundEmail) {
//                 console.log("Email not found in CSV");
//             }

//             resolve(accessGranted); 
//         })
//         .catch(error => {
//             console.error("Error:", error);
//             reject(error); 
//         });
//     });
// }

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
