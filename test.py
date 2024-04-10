from taskweaver.app.app import TaskWeaverApp
import pandas as pd
# This is the folder that contains the taskweaver_config.json file and not the repo root. Defaults to "./project/"
app_dir = "./project/"

from fastapi import FastAPI

app = FastAPI()

from pydantic import BaseModel

class UserCreate(BaseModel):
        json_table: 
        username: str

@app.post("/create_user/")
async def clinical_inputs_table_5_1():
    
    app = TaskWeaverApp(app_dir=app_dir)
    session = app.get_session()
    user_query = f"""
    Perform following steps sequentially, only once you finish one step you can go to next step,
    Step 1: load the file C:\\Users\\2113196\\Downloads\\Clinical inputs_Section 5.1.xlsx skipping 4 rows as df1
    Step 2: Sum up ‘Patients treated’ column
    Step 3: Replace Placeholder 1 in below paragraph with sum of patients treated from previous step
    Paragraph: Cumulatively '''Placeholder 1''' patients have been exposed to Sevelamer in Cognizant-sponsored investigational 
    clinical trials.
    """
    response_round = session.send_message(user_query)
    response_dict = response_round.to_dict()

    # Extract and print the message from the last post in the post_list
    # Assuming the last post is the response from TaskWeaver
    last_post_message = response_dict['post_list'][-1]['message']
    print(last_post_message)

    # load the file excel file converted from JSON df2 ***** Python code for that here

    user_query = f"""
    Step 4: *************
    Step 5: Filter on column ‘Formulation’=‘Tablet’ on df1
    Step 6: Now, add values of column ‘Patients treated’ where column ‘Status’ is ‘Completed’
    Step 7: df2.iat[0,1]= Sum from previous step
    Step 8: Now, add values of column ‘Patients treated’ where column ‘Status’ is ‘Ongoing’
    Step 9: df2.iat[1,1]= Sum from previous step
    Step 10: save same excel file, don’t create a copy
    """
    response_round = session.send_message(user_query)

    app.stop()

    user_query = f"""
    Perform following steps sequentially, only once you finish one step you can go to next step,
    Step 1: load the file C:\\Users\\2113196\\Downloads\\Clinical inputs_Section 5.1.xlsx skipping 4 rows as df1
    Step 2: ----------------------------
    Step 3: Filter on column ‘Formulation’=‘Powder for oral suspension’ on df1
    Step 4: Now, add values of column ‘Patients treated’ where column ‘Status’ is ‘Completed’
    Step 5: df2.iat[0,2]= Sum from previous step
    Step 6: Now, add values of column ‘Patients treated’ where column ‘Status’ is ‘Ongoing’
    Step 7: df2.iat[1,2]= Sum from previous step
    Step 8: save same excel file, don’t create a copy
    """
    response_round = session.send_message(user_query)

    app.stop()
    
    user_query = f"""
    Perform following steps sequentially, only once you finish one step you can go to next step,
    Step 1: load the file C:\\Users\\2113196\\Downloads\\Clinical inputs_Section 5.1.xlsx skipping 4 rows as df1
    Step 2: -----------------------------
    Step 3: Filter on column ‘Formulation’=‘Tablet and Powder for oral suspension’ on df1
    Step 4: Now, add values of column ‘Patients treated’ where column ‘Status’ is ‘Completed’
    Step 5: df2.iat[0,3]= Sum from previous step
    Step 6: Now, add values of column ‘Patients treated’ where column ‘Status’ is ‘Ongoing’
    Step 7: df2.iat[1,3]= Sum from previous step
    Step 8: save same excel file, don’t create a copy
    """
    response_round = session.send_message(user_query)

    app.stop()

    user_query = f"""
    Perform following steps sequentially, only once you finish one step you can go to next step,
    Step 1: Create table or load excel as df2
    Step 2: df2.iat[0,4]=Sum of that row
    Step 3: df2.iat[1,4]=Sum of that row
    Step 4: df2.iat[2,1]=Sum of that column
    Step 5: df2.iat[2,2]=Sum of that column
    Step 6: df2.iat[2,3]=Sum of that column
    Step 7: df2.iat[2,4]=Sum of that column
    Step 8: save same excel file, don’t create a copy
    """
    response_round = session.send_message(user_query)


    path_excel = '------------------'
    df = pd.read_excel(path_excel, engine='openpyxl')

    # Convert DataFrame to JSON
    json_data = df.to_json(orient='records', indent=4)
    print(json_data)



    app.stop()

clinical_inputs_table_5_1()


# app = TaskWeaverApp(app_dir=app_dir)
# session = app.get_session()

# app.__getstate__()

# LL_reporting_location = input("Enter location for LL_reporting.xlsx file: ")

# user_query = f"""Perform following steps sequentially, only once you finish one step you can go to next step,
# Step 1: load the file {LL_reporting_location} as df1
# Step 2: load the file C:\\Users\\2113196\\Downloads\\LineListingTable.xlsx as df2

# Step 3: Filter df1 on ‘Report Type’ = ‘Clinical Trial’, ‘Medically confirmed case’ = ‘Yes’ and ‘Seriousness (Event)’ = ‘Serious’ 
# Step 4: df2.iat[1,1]= Count of rows from filtered df
 
# Step 5: Filter df1 on ‘Report Type’ = ‘Clinical Trial’, ‘Medically confirmed case’ = ‘No’ and ‘Seriousness (Event)’ = ‘Serious’
# Step 6: df2.iat[1,2]= Count of rows from filtered df

# Step 7: Filter df1 on ‘Report Type’ = ‘Clinical Trial’, ‘Medically confirmed case’ = ‘Yes’ and ‘Seriousness (Event)’ = ‘Non-Serious’
# Step 8: df2.iat[1,3]= Count of rows from filtered df

# Step 9: Filter df1 on ‘Report Type’ = ‘Clinical Trial’, ‘Medically confirmed case’ = ‘No’ and ‘Seriousness (Event)’ = ‘Non-Serious’
# Step 10: df2.iat[1,4]= Count of rows from filtered df
# Step 11: Save the same excel file, don't create a copy

# """
# response_round = session.send_message(user_query)
# print(response_round.to_dict())

# app.stop()

# app = TaskWeaverApp(app_dir=app_dir)
# session = app.get_session()

# user_query = """
# Perform following steps sequentially, only once you finish one step you can go to next step,
# Step 1: load the file C:\\Users\\2113196\\Downloads\\LL_reporting.xlsx as df1
# Step 2: load the file C:\\Users\\2113196\\Downloads\\LineListingTable.xlsx as df2

# Step 3: Filter df1 on ‘Report Type’ = ‘Literature’, ‘Medically confirmed case’ = ‘Yes’ and ‘Seriousness (Event)’ = ‘Serious’
# Step 4: df2.iat[2,1]= Count of rows from filtered df

# Step 5: Filter df1 on ‘Report Type’ = ‘Literature’, ‘Medically confirmed case’ = ‘No’ and ‘Seriousness (Event)’ = ‘Serious’
# Step 6: df2.iat[2,2]= Count of rows from filtered df

# Step 7: Filter df1 on ‘Report Type’ = ‘Literature’, ‘Medically confirmed case’ = ‘Yes’ and ‘Seriousness (Event)’ = ‘Non-Serious’
# Step 8: df2.iat[2,3]= Count of rows from filtered df

# Step 9: Filter df1 on ‘Report Type’ = ‘Literature’, ‘Medically confirmed case’ = ‘No’ and ‘Seriousness (Event)’ = ‘Non-Serious’
# Step 10: df2.iat[2,4]= Count of rows from filtered df
# Step 11: Save the same excel file, don't create a copy
# """
# response_round = session.send_message(user_query)
# print(response_round.to_dict())

# app.stop()

# app = TaskWeaverApp(app_dir=app_dir)
# session = app.get_session()

# user_query = """
# Perform following steps sequentially, only once you finish one step you can go to next step,
# Step 1: load the file C:\\Users\\2113196\\Downloads\\LL_reporting.xlsx as df1
# Step 2: load the file C:\\Users\\2113196\\Downloads\\LineListingTable.xlsx as df2

# Step 19: Filter df1 on ‘Report Type’ = ‘Spontaneous Report’, ‘Medically confirmed case’ = ‘Yes’ and ‘Seriousness (Event)’ = ‘Serious’
# Step 4: df2.iat[3,1]= Count of rows from filtered df

# Step 21: Filter df1 on ‘Report Type’ = ‘Spontaneous Report’, ‘Medically confirmed case’ = ‘No’ and ‘Seriousness (Event)’ = ‘Serious’
# Step 6: df2.iat[3,2]= Count of rows from filtered df

# Step 23: Filter df1 on ‘Report Type’ = ‘Spontaneous Report’, ‘Medically confirmed case’ = ‘Yes’ and ‘Seriousness (Event)’ = ‘Non-Serious’
# Step 8: df2.iat[3,3]= Count of rows from filtered df

# Step 25: Filter df1 on ‘Report Type’ = ‘Spontaneous Report’, ‘Medically confirmed case’ = ‘No’ and ‘Seriousness (Event)’ = ‘Non-Serious’
# Step 10: df2.iat[3,4]= Count of rows from filtered df
# Step 11: Save the same excel file, don't create a copy
# """
# response_round = session.send_message(user_query)
# print(response_round.to_dict())

# app.stop()

# app = TaskWeaverApp(app_dir=app_dir)
# session = app.get_session()

# user_query = """
# Perform following steps sequentially, only once you finish one step you can go to next step,
# Step 1: load the file C:\\Users\\2113196\\Downloads\\LL_reporting.xlsx as df1
# Step 2: load the file C:\\Users\\2113196\\Downloads\\LineListingTable.xlsx as df2

# Step 27: Filter df1 on ‘Report Type’ = ‘Post Market Survey’, ‘Medically confirmed case’ = ‘Yes’ and ‘Seriousness (Event)’ = ‘Serious’
# Step 4: df2.iat[4,1]= Count of rows from filtered df

# Step 29: Filter df1 on ‘Report Type’ = ‘Post Market Survey’, ‘Medically confirmed case’ = ‘No’ and ‘Seriousness (Event)’ = ‘Serious’
# Step 6: df2.iat[4,2]= Count of rows from filtered df

# Step 31: Filter df1 on ‘Report Type’ = ‘Post Market Survey’, ‘Medically confirmed case’ = ‘Yes’ and ‘Seriousness (Event)’ = ‘Non-Serious’
# Step 8: df2.iat[4,3]= Count of rows from filtered df

# Step 33: Filter df1 on ‘Report Type’ = ‘Post Market Survey’, ‘Medically confirmed case’ = ‘No’ and ‘Seriousness (Event)’ = ‘Non-Serious’
# Step 10: df2.iat[4,4]= Count of rows from filtered df
# Step 11: Save the same excel file, don't create a copy
# """
# response_round = session.send_message(user_query)
# print(response_round.to_dict())

