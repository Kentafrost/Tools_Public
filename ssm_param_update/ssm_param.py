import os
import boto3
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
import logging

gsheet_cred = "G:\\My Drive\\Tool\\credentials.json"

# gsheet authentication
def authorize_gsheet():
    try:
        scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        credentials = Credentials.from_service_account_file(
            gsheet_cred, scopes=scope
        )
        
        gspread_client = gspread.authorize(credentials)
        return gspread_client
    except Exception as e:
        # stop the program if authentication fails
        raise SystemExit("Googleの認証処理でエラーが発生しました。{}".format(e))


# check all ssm parameters in SSM Parameter Store and list them
def put_ssm_parameter(ssm_client, df):

    before_dict = {}
    current_dict = {}
    create_log = []
    
    # retrieve all parameters from ssm parameters and store in list
    all_parameters = []
    next_token = None

    while True:
        if next_token:
            response = ssm_client.describe_parameters(NextToken=next_token)
        else:
            response = ssm_client.describe_parameters()
        all_parameters.extend(response['Parameters'])
        next_token = response.get('NextToken')
        if not next_token:
            break

    current_param_names = [param['Name'] for param in all_parameters]

    # store all parameters in current SSM Parameter Store
    print("====================")
    print("Current SSM Parameters names:")
    print(current_param_names)
    print("====================")
    
    # firstly, check if parameters already exist in SSM Parameter Store
    for index, row in df.iterrows():
        # skip the first row which is header
        if index == 0:
            continue
        
        csv_param_name = row[0]
        csv_param_value = row[1]
        csv_param_type = row[2]

        try:
            # check if the parameter already exists, if yes, update the value
            if csv_param_name in current_param_names:
                print(f"Parameter {csv_param_name} already exists.")

                # get the current value of the parameter if it exists
                response = ssm_client.get_parameter(
                    Name=csv_param_name, WithDecryption=True
                )
                
                # update the parameter value
                try:
                    ssm_client.put_parameter(
                        Name=csv_param_name,
                        Value=csv_param_value,
                        Type=csv_param_type,
                        Overwrite=True
                    )
                    # list up parameters key, and value after updating
                    current_dict[csv_param_name] = csv_param_value
                    logging.info(f"Updated parameter {csv_param_name} to {csv_param_value}")
                    
                    status = "Parameter Updated"
                        
                except Exception as e:
                    print(f"Error updating parameter {csv_param_name}: {e}")
                    logging.error(f"Error updating parameter {csv_param_name}: {e}")
                    pass

            else:
                # if parameter not in before_dict, then just update the value
                before_dict[csv_param_name] = "None"

                # create a new parameter
                print(f"Creating new parameter {csv_param_name} with value {csv_param_value}")
                
                ssm_client.put_parameter(
                    Name=csv_param_name,
                    Value=csv_param_value,
                    Type=csv_param_type,
                    Overwrite=False
                )
                # list up parameters key, and value after updating
                current_dict[csv_param_name] = csv_param_value
                print(f"Parameter {csv_param_name} does not exist. Created new parameter.")
                
                status = "New Parameter Created"
                
            create_log.append({
                "Timestamp": pd.Timestamp.now(),
                "Parameter Name": csv_param_name,
                "Parameter Value": csv_param_value,
                "Parameter Type": csv_param_type,
                "Status": status
            })

        except Exception as e:
            print(f"Error checking SSM parameter: {e}")
            print("Please check the parameter name and value.")
            
            create_log.append({
                "Timestamp": pd.Timestamp.now(),
                "Parameter Name": csv_param_name,
                "Parameter Value": csv_param_value,
                "Parameter Type": csv_param_type,
                "Status": "Error - " + str(e)
            })

    return create_log


def delete_ssm_parameter(ssm_client, df):
    
    delete_log = []
    
    for index, row in df.iterrows():
        # skip the first row which is header
        if index == 0:
            continue
        
        csv_param_name = row[0]
        csv_param_value = row[1]
        csv_param_type = row[2]
        
        ssm_client.delete_parameter(Name=csv_param_name)
        delete_log.append({
            "Timestamp": pd.Timestamp.now(),
            "Parameter Name": csv_param_name,
            "Parameter Value": csv_param_value,
            "Parameter Type": csv_param_type,
            "Status": "Deleted"
        })
        
    return delete_log


if __name__ == "__main__":
    
    ssm_client = boto3.client('ssm', region_name='ap-southeast-2')
    # current directory
    current_dir = os.getcwd()
    
    # get all SSM parameters from gsheet
    gc = authorize_gsheet()
    workbook = gc.open("SSM_param")
    
    sheet = workbook.worksheet("SSM_Add")
    data = sheet.get_all_values()
    df = pd.DataFrame(data)
    create_log = put_ssm_parameter(ssm_client, df)
    
    sheet = workbook.worksheet("SSM_Delete")
    data = sheet.get_all_values()
    df = pd.DataFrame(data)
    delete_log = delete_ssm_parameter(ssm_client, df)

    all_parameters = []
    next_token = None

    while True:
        if next_token:
            response = ssm_client.describe_parameters(NextToken=next_token)
        else:
            response = ssm_client.describe_parameters()
        all_parameters.extend(response['Parameters'])
        next_token = response.get('NextToken')
        if not next_token:
            break
    
    df_param = pd.DataFrame(all_parameters)
    df_param = df_param.sort_values(by='Name')

    this_script_dir = os.path.dirname(os.path.abspath(__file__))
    df_param.to_csv(f"{this_script_dir}//ssm_parameters.csv", index=False)

    # save create and delete logs to CSV files
    create_log_df = pd.DataFrame(create_log)
    create_log_df.to_csv(f"{this_script_dir}//ssm_create_log.csv", index=False)
    
    delete_log_df = pd.DataFrame(delete_log)
    delete_log_df.to_csv(f"{this_script_dir}//ssm_delete_log.csv", index=False)

    print("All SSM parameters have been updated successfully.")