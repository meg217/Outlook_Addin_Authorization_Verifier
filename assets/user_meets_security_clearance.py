import csv # This is needed for the user_meets_security_clearance(...) function for csv parsing

# This function checks will return if the user's clearance meets requirements i.e. a user's TOP SECRET clearance will guarantee access may despite document classification being CONFIDENTIAL
def user_meets_security_clearance(file_path, document_classification, email):
    try:
        # Open the CSV file
        with open(file_path, mode='r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            
            # Iterate through each row in the CSV
            # Edge Case to discuss: What if user has multiple records with varying clearances?
            for row in csv_reader:
                # Checking to see if record is the user
                if row['Email'] == email:
                    user_clearance = row['Classification']
                    
                    if (can_user_access_confidential(document_classification,user_clearance)):
                        return True
                    if (can_user_access_secret(document_classification,user_clearance)):
                        return True
                    if (can_user_access_top_secret(document_classification,user_clearance)):
                        return True
                    
                    # User does not have enough clearance to access the document    
                    #print(email,"cannot access document...")
                    
    except Exception as e:
        print(f"An error occurred: {e}")
    
    return False

def can_user_access_confidential(document_classification, user_clearance):
    if (document_classification.strip().lower() == "confidential"):
        if (user_clearance.strip().lower() == "confidential" or user_clearance.strip().lower() == "secret" or user_clearance.strip().lower() == "top secret"):
            return True
    return False
def can_user_access_secret(document_classification, user_clearance):
    if (document_classification.strip().lower() == "secret"):
        if (user_clearance.strip().lower() == "secret" or user_clearance.strip().lower() == "top secret"):
            return True
    return False
def can_user_access_top_secret(document_classification, user_clearance):
    if (document_classification.strip().lower() == "top secret"):
        if (user_clearance.strip().lower() == "top secret"):
            return True
    return False

# Example usage
file_path = 'users.csv'  # Adjust as needed
document_classification = 'secret'
email = 'johndoe@yahoo.tx.gov'
result = user_meets_security_clearance(file_path, document_classification, email)
print(result)
