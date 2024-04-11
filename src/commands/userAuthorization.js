const fs = require('fs');
const csv = require('csv-parser');

// This function checks if the user's clearance meets requirements
function userMeetsSecurityClearance(filePath, documentClassification, email) {
    return new Promise((resolve, reject) => {
        let accessGranted = false;

        fs.createReadStream(filePath)
            .pipe(csv())
            .on('data', (row) => {
                if (row.Email === email) {
                    const userClearance = row.Classification;

                    if (canUserAccess(documentClassification, userClearance)) {
                        accessGranted = true;
                    }
                }
            })
            .on('end', () => {
                resolve(accessGranted);
            })
            .on('error', (error) => {
                reject(error);
            });
    });
}

function canUserAccess(documentClassification, userClearance) {
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
