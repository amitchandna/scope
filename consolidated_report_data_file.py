import psycopg2
import pandas as pd

# Connect to the PostgreSQL database
conn = psycopg2.connect(**params)
# Create a new cursor
cur = conn.cursor()


# A function that takes in a PostgreSQL query and outputs a pandas dataframe
def create_pandas_table(sql_query, database=conn):
    table = pd.read_sql_query(sql_query, database)
    return table


# Utilize the create_pandas_table  function to create a Pandas data frame
# Store the data as a variable
project_info = create_pandas_table("""SELECT id AS project_id, name, deal_number FROM public.projects_project
                                   WHERE deal_number!= 9999 AND deal_number IS NOT NULL""")
product_tools = create_pandas_table("""SELECT products_tool.id AS tool_id, products_tool.version_major, products_tool.type_id, products_tool.version_minor, products_tool.version_revision, products_tooltype.description
                                    FROM public.products_tool 
                                    RIGHT JOIN public.products_tooltype ON products_tooltype.id = products_tool.type_id 
                                    WHERE (products_tool.version_major = 2 or products_tool.version_major = 4) 
                                    AND (products_tooltype.description = 'SCOPE Basic' OR products_tooltype.description = 'SCOPE Pro') """)
assessments = create_pandas_table("""SELECT id as assessment_id, score, project_id, assessor_full_name, assessor_user_id, report_status 
                                  FROM public.assessments_assessment where report_status='final'""")
assessments_section_response = create_pandas_table("""SELECT 
                                                   id AS section_response_id, 
                                                   score, 
                                                   section_title, 
                                                   section_display_position, 
                                                   _assessment_id AS assessment_id 
                                                   FROM public.assessments_sectionresponse
                                                   WHERE assessments_sectionresponse._assessment_id in
                                                    (SELECT assessments_assessmentl.id FROM public.assessments_assessment WHERE assessments_assessment.report_status='final')""")
account_user = create_pandas_table("""SELECT id as user_id, 
                                   country, 
                                   global_region, 
                                   region
                                   FROM public.accounts_user""")
producing_organization_details = create_pandas_table("""SELECT id, 
                                   number_of_female_executives, 
                                   number_of_male_executives, 
                                   number_of_female_non_executives, 
                                   number_of_male_non_executives, 
                                   number_of_female_members, 
                                   number_of_male_members, 
                                   number_of_female_full_time_employees, 
                                   number_of_male_full_time_employees, 
                                   number_of_full_time_employees, 
                                   _assessment_id AS assessment_id 
                                   FROM public.assessments_producingorganizationdetails
                                   WHERE _assessment_id IS NOT NULL""")

# Create a Pandas Excel writer using XlsxWriter as the engine.
with pd.ExcelWriter('outout.xlsx', engine='xlsxwriter') as writer:
    # Write each dataframe to a different worksheet.
    project_info.to_excel(writer, sheet_name='Project Info', index = False)
    product_tools.to_excel(writer, sheet_name='Project Tools', index = False)
    assessments.to_excel(writer, sheet_name='Assessments', index = False)
    assessments_section_response.to_excel(writer, sheet_name='Assessments Section Response', index = False)
    account_user.to_excel(writer, sheet_name = 'Account User', index = False)
    producing_organization_details.to_excel(writer, sheet_name = 'Producing Organization Details', index = False)

# Close the cursor and connection to so the server can allocate
# bandwidth to other requests
cur.close()
conn.close()