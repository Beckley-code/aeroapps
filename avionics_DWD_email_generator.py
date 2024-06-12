############################################################################
#  Description:
#
#
#
#
#
#  Author: Matt Sunday (2194176)
#
#  Revision History:
#
#
#
#  Known Limitations:
#    - Shayla STOMPRO has her last name all capitalized in one of the source files
#      so all the dictionaries also have her name capitalized - no loss in function,
#      it just looks odd when reading through the code
#    - This script only supports emailing for ETAC events
#
#
############################################################################
##
##pip install wheel
import pandas as pd
import numpy
from datetime import datetime
from datetime import timedelta
from datetime import date
import operator
import sys
import math
import warnings
from pprint import pprint
from pathlib import Path
import win32com.client as Client

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

class email:
    def __init__(self, _to, _cc, _subject):
        self.body = ""
        self.subject = _subject
        self.to = _to
        self.cc = _cc

    def create_body(self, _data_list, _dwd_list_for_email):
        body_str = "Below is your"

        for data, title in zip(_data_list, _dwd_list_for_email):
            pass

    def send_email(self):
        outlook = Client.Dispatch('Outlook.Application')
        message = outlook.CreateItem(0)
        message.To = self.to
        message.CC = self.cc
        message.BCC = ""
        message.Subject = self.subject
        message.HTMLBody = self.body
        message.Display()
        #message.Send()


def load_spreadsheet(_filename, _sheet="", skip=0):
    print("--> Opening File: {}".format(_filename))
    raw_data = pd.read_excel(_filename, _sheet, skiprows=skip)
    return raw_data


def load_group_code_data():
    print("-> Loading Group Code Data from Server")
    _filename = "\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\Mapping Data\\managers_to_group_codes.xlsx"
    _sheet = "manager"
    raw_data = load_spreadsheet(_filename, _sheet, skip=0)
    
    managers = raw_data['Supervisor'].unique()  # list of unique manager names

    _manager_to_gc_dict = {}

    for manager in managers:
        group_codes = list(raw_data[raw_data['Supervisor'] == manager]['GC'])  # list of all group codes matching manager, includes duplicates
        group_codes = list(set(group_codes))  # removes duplicates
        _manager_to_gc_dict[manager] = group_codes  # assign group code list as the value to the manager key

    return _manager_to_gc_dict


def filter_etac_data(_data, _data_type, _group_codes):
    """
    This function is for filtering the ETAC data only
    
    """

    new_tmp_data_filter = _data['GC/OBS'].isin(_group_codes)  # creates filter by group code's that are applicable to the manager in question
    new_tmp_data = _data[new_tmp_data_filter]  # actually filter the DataFrame
    new_tmp_data = new_tmp_data[new_tmp_data['STATUS'] != 'On Time']  # filter out completes
    new_tmp_data.sort_values(by=['DATE DUE'], ascending=False)  # sort by most neart-term due

    if "CG12" in _data_type:
        new_tmp_data.columns = ['FUNCTION', 'COMMODITY', 'DATE DUE', 'DATE ECD', 'DATE COMP', 'STATUS', 'COORD GROUP',
                                'CG', 'TOR', 'PRIORITY', 'CHANGE PRIORITY', 'RELREC', 'TASK OWNER',
                                'COORDINATOR', 'CUSTOMER', 'MODEL', 'LINE', 'CG/OBS', 'WORK NUMBER',
                                'REMARKS', 'DESCRIPTION', 'TEAM', 'ETAC LOAD DATE', 'MANAGER BEMS UNION',
                                'OB DUE', 'OB ECD', 'OB COMP']  # this is needed because the export has two columns named 'REMARKS'
    elif ("CG48" or "CG58") in _data_type and "CG3875" not in _data_type:
        new_tmp_data.columns = ['FUNCTION', 'COMMODITY', 'DATE DUE', 'DATE ECD', 'DATE COMP', 'STATUS', 'COORD GROUP',
                                'CG', 'TOR', 'PRIORITY', 'CHANGE PRIORITY', 'RELREC', 'TASK OWNER',
                                'COORDINATOR', 'CUSTOMER', 'MODEL', 'LINE', 'CG/OBS', 'WORK NUMBER',
                                'REMARKS', 'DESCRIPTION', 'TEAM', 'ETAC LOAD DATE', 'MANAGER BEMS UNION',
                                'ENGR REL ECD COMMENT']  # this is needed because the export has two columns named 'REMARKS'
    elif "CG3875" in _data_type:
        new_tmp_data.columns = ['FUNCTION', 'COMMODITY', 'DATE DUE', 'DATE ECD', 'DATE COMP', 'STATUS',
                                'CG', 'TOR', 'PRIORITY', 'CHANGE PRIORITY', 'RELREC', 'TASK OWNER',
                                'COORDINATOR', 'CUSTOMER', 'MODEL', 'LINE', 'CG/OBS', 
                                'DESCRIPTION', 'REMARKS', 'WORK NUMBER', 'TEAM', 'ETAC LOAD DATE', 
                                'MANAGER BEMS UNION', 'OB DUE', 'OB ECD', 'OB COMP']  # this is needed because the export has two columns named 'REMARKS'

    return new_tmp_data


def add_etac_table(_data):
    """

    """
    
    tmp_str = "<table border='1'; style='border-collapse:collapse'; align='center'; width=100%>"
    tmp_str = tmp_str + "<tr><th style='text-align: center'>Due</th>\
                             <th style='text-align: center'>ECD</th>\
                             <th style='text-align: center'>TOR</th>\
                             <th style='text-align: center'>Link</th>\
                             <th style='text-align: center'>Owner</th>\
                             <th style='text-align: center'>Model</th>\
                             <th style='text-align: center'>L/N</th>\
                             <th style='text-align: center'>Work Number</th>\
                             <th style='text-align: center'>Notes</th>\
                             <th style='text-align: center'>Description</th>\
                             </tr>"

    for due_date, ecd_date, status, tor, relrec, owner, model, LN, work_number, remarks, description in zip(_data['DATE DUE'], _data['DATE ECD'],_data['STATUS'],_data['TOR'],_data['RELREC'],_data['TASK OWNER'],_data['MODEL'],_data['LINE'],_data['WORK NUMBER'],_data['REMARKS'],_data['DESCRIPTION']):
        time_yesterday = datetime.now() - timedelta(days=1)  # account for ETAC due date being midnight and python script being current time
        if due_date < time_yesterday: _style = "background-color:#FFCCCB"
        else: _style = "background-color:#FFFFFF"
        
        #if status == "Delinquent": _style = "background-color:#FFCCCB"
        #else: _style = "background-color:#FFFFFF"

        if math.isnan(LN): LN = "N/A"
        else: LN = int(LN)

        tmp_str = tmp_str + "<tr style=" + _style + ">"
        tmp_str = tmp_str + "<td style='text-align: center; width: 45px'>{}</td>".format(str(due_date)[5:10])
        tmp_str = tmp_str + "<td style='text-align: center; width: 45px'>{}</td>".format(str(ecd_date)[5:10])
        tmp_str = tmp_str + "<td style='text-align: center'>{}</td>".format(tor)
        tmp_str = tmp_str + "<td style='text-align: center'>{}</td>".format(relrec)
        tmp_str = tmp_str + "<td style='text-align: center'>{}</td>".format(owner)
        tmp_str = tmp_str + "<td style='text-align: center'>{}</td>".format(model)
        tmp_str = tmp_str + "<td style='text-align: center'>{}</td>".format(LN)
        tmp_str = tmp_str + "<td style='text-align: center'>{}</td>".format(work_number)
        tmp_str = tmp_str + "<td>{}</td>".format(remarks)
        tmp_str = tmp_str + "<td style='text-align: center'>{}</td>".format(description)
        tmp_str = tmp_str + "</tr>"        

    tmp_str = tmp_str + "</table><br><br>"

    return tmp_str

    
def add_to_email_body(_emailobj, _data_type, _data, _name=""):
    """

    """

    if "current" in _data_type:  week_str = "This Week" # define the string to put in the email for the header text 
    elif "next" in _data_type: week_str = "Next Week"

    if "CG" in _data_type:  # ETAC items
        if "CG12" in _data_type: data_type_str = "{} CG 1 and 2 - ".format(_name)
        elif "CG48" in _data_type: data_type_str = "{} CG 48 - ".format(_name)
        elif "CG58" in _data_type: data_type_str = "{} CG 58 - ".format(_name)
        elif "CG3875" in _data_type: data_type_str = "{} CG 38 and 75 - ".format(_name)

        header_str = data_type_str + week_str
        #scheduled_str = str(len(_data[_data['STATUS'] == 'Scheduled']))
        #delinquent_str = str(len(_data[_data['STATUS'] == 'Delinquent']))
        time_yesterday = datetime.now() - timedelta(days=1)
        delinquent_str = str(len(_data[_data['DATE DUE'] < time_yesterday]))
        scheduled_str = str(len(_data[_data['DATE DUE'] >= time_yesterday]))

        # Add on to the body of the email
        _emailobj.body = _emailobj.body + "<div>"
        _emailobj.body = _emailobj.body + "<h4>" + header_str + "</h4>"
        _emailobj.body = _emailobj.body + "- Delinquent: " + delinquent_str
        _emailobj.body = _emailobj.body + "<br>- Scheduled: " + scheduled_str + "<br><br>"
        _emailobj.body = _emailobj.body + add_etac_table(_data)
        _emailobj.body = _emailobj.body + "</div>"
        #################################

        #print(_emailobj.body)

    return


def add_footer_to_body(_emailobj):
    """

    """

    time_now = datetime.now()
    time_str = str(time_now)

    _emailobj.body = _emailobj.body + "<br>Note: This was generated on " + time_str + " and is only as good as the BEDAT extracts \
                                  that are stored at the following location: \\\\nw\data\AVI\RTB Weekly Metrics for Managers<br><br>\
                                  For any real-time data, please visit the following website, or BEDAT: https://my.boeing.com/myb/myportal/BCA/bcaairsys/Avionics/Meetings/"

    return


if __name__ == "__main__":

    ############################################################################
    #This section defines all the global variables for filenames, senior to 
    #manager mappings, etc
    ############################################################################
    file_data = {"CG12_current": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG12_currentweek", 'DWD ETAC Detail_1'],  # This dictionary is all the source files for the content to import, and the associated sheet names
                 "CG12_next": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG12_nextweek", 'DWD ETAC Detail_1'],
                 "CG48_current": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG48_currentweek", 'DWD ETAC Detail_1'],
                 "CG48_next": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG48_nextweek", 'DWD ETAC Detail_1'],
                 "CG58_current": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG58_currentweek", 'DWD ETAC Detail_1'],
                 "CG58_next": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG58_nextweek", 'DWD ETAC Detail_1'],
                 "CG3875_current": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG3875_currentweek", 'DWD ETAC Detail_1'],
                 "CG3875_next": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG3875_nextweek", 'DWD ETAC Detail_1'],
                 "MM_current": ["", ],
                 "MM_next": ["", ],
                 "SCN": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\SCN_Backlog.xlsx", ],
                 "PNNC": ["", ],
                 "BCAB": ["", ]}

    file_data = {"CG12_current": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG12_currentweek.xlsx", 'DWD ETAC Detail_1'],
                 "CG12_next": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG12_nextweek.xlsx", 'DWD ETAC Detail_1'],
                 "CG48_current": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG48_currentweek.xlsx", 'DWD ETAC Detail_1'],
                 "CG48_next": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG48_nextweek.xlsx", 'DWD ETAC Detail_1'],
                 "CG58_current": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG58_currentweek.xlsx", 'DWD ETAC Detail_1'],
                 "CG58_next": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG58_nextweek.xlsx", 'DWD ETAC Detail_1'],
                 "CG3875_current": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG3875_currentweek.xlsx", 'DWD ETAC Detail_1'],
                 "CG3875_next": ["\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\CG3875_nextweek.xlsx", 'DWD ETAC Detail_1']}  # TEMPORARY just to run the script on a single input file for testing

    dwd_list_for_email = ["CG 1 and 2 - Current Week",
                          "CG 1 and 2 - Next Week", 
                          "CG 48 - Current Week", 
                          "CG 48 - Next Week", 
                          "CG 58 - Current Week",
                          "CG 58 - Next Week",
                          "CG 38 and 75 - Current Week", 
                          "CG 38 and 75 - Next Week",
                          "Major / Minors - Current Week",
                          "Major / Minors - Next Week",
                          "SCNs",
                          "pNNCs",
                          "BCAB Action Items"]

    director = 'Sunday'  # assigned a variable for ease of scaling / changing

    senior_to_manager_dict = {'Mai': ['Chand', 'Miller', 'Bourgeois'],
                              'Carlson':['Bhowmick', 'Laxton', 'Taylor', 'Vonjouanne'],
                              'Jayaram': ['Jones', 'Shalabi'],
                              'Strong': ['Bement', 'Goyer', 'Kausar', 'Westerlund', 'Williams', 'Wilkins'],
                              'Duggal': ['Alnoor', 'Noorfeshan', 'STOMPRO', 'McGuire'],
                              'Caballero': ['Kahle'],
                              'McClure': ['Prieto', 'York'],
                              'Haq': ['Awan', 'Saldana', 'Quedado']}

    manager_to_email_dict = {'Mai': 'rochelle.g.mai@boeing.com',
                             'Chand': 'uma.chand@boeing.com', 
                             'Miller': 'colleen.m.miller2@boeing.com', 
                             'Bourgeois': 'brian.d.bourgeois@boeing.com',
                             'Carlson': 'stephanie.m.carlson@boeing.com',
                             'Bhowmick': 'rahul.bhowmick@boeing.com', 
                             'Laxton': 'caitlin.m.laxton@boeing.com', 
                             'Taylor': 'emilie.m.taylor@boeing.com', 
                             'Vonjouanne': 'henry.v.vonjouanne@boeing.com',
                             'Jayaram': 'sanjay.jayaram@boeing.com',
                             'Jones': 'katherine.jones2@boeing.com', 
                             'Shalabi': 'diana.a.shalabi@boeing.com',
                             'Strong': 'rachelle.l.strong@boeing.com',
                             'Bement': 'nathan.j.bement@boeing.com', 
                             'Goyer': 'erin.a.goyer@boeing.com', 
                             'Kausar': 'saleh.m.kausar@boeing.com', 
                             'Westerlund': 'larry.d.westerlund@boeing.com', 
                             'Williams': 'samuel.f.williams@boeing.com', 
                             'Wilkins': 'colin.j.wilkins@boeing.com',
                             'Duggal': 'rohit.duggal@boeing.com',
                             'Alnoor': 'zina.alnoor@boeing.com', 
                             'Noorfeshan': 'nader.noorfeshan@boeing.com', 
                             'STOMPRO': 'shayla.a.stompro@boeing.com', 
                             'McGuire': 'madeline.j.mcguire@boeing.com',
                             'Caballero': 'jorge.e.caballero2@boeing.com',
                             'Kahle': 'william.p.kahle@boeing.com',
                             'McClure': 'casey.a.mcclure@boeing.com', 
                             'Prieto': 'cesar.h.prieto@boeing.com', 
                             'York': 'casey.a.york@boeing.com',
                             'Sunday': 'matthew.m.sunday@boeing.com',
                             'Awan': 'saira.s.awan@boeing.com',
                             'Haq': 'hasan.haq@boeing.com',
                             'Quedado': 'arjaysteven.r.quedado@boeing.com',
                             'Saldana': 'faith.p.saldana@boeing.com'}

    manager_to_gc_dict = load_group_code_data()  # loads a dictionary of manager last names as keys and a list of Group Codes as the value
    #TODO: Chand is not loading from the file...oithers may also not be loading

    #employee_to_manager_dict  # TODO: needed for non-ETAC events
    ############################################################################



    ############################################################################
    # This section loops through all the senior_to_manager_dict and creates an
    # email class object for each manager and each senior manager
    ############################################################################
    director_emailobj_dict = {}  # director dictionary and could be used to add others like Jake, Eli, etc
    senior_emailobj_dict = {}
    manager_emailobj_dict = {}
    current_time = datetime.now()  # current time object
    todays_date = current_time.strftime('%m/%d/%Y')  # format current time to MM/DD/YYYY

    for senior in senior_to_manager_dict.keys():  # loop through each senior
        senior_emailobj_dict[senior] = email(manager_to_email_dict[senior], manager_to_email_dict[director], "{} Do What's Due Update for {}".format(todays_date, senior))  # email class for the manager
        for manager in senior_to_manager_dict[senior]:  # loop through each manager for the senior 
            manager_emailobj_dict[manager] = email(manager_to_email_dict[manager], manager_to_email_dict[senior], "{} Do What's Due Update for {}".format(todays_date, manager))  # email class for the senior managers
    director_emailobj_dict[director] = email(manager_to_email_dict[director], "", "{} Do What's Due Update for {}".format(todays_date, director))  # email class for the director
    ############################################################################



    ############################################################################
    # This section loops through all files and imports the data into pandas
    # dataframes, and stores these dataframes in the dictionary file_data
    ############################################################################
    print("\n-> Loading all Data from the Server")
    for data_type in file_data.keys():
        # data_type = the name of the type of data CG12_current, CG12_next, MM_current, SCN, etc
        filename = file_data[data_type][0]  # filename of the source data
        sheet_name = file_data[data_type][1]  # sheet name of the source data
        
        if Path(filename).is_file():
            file_data[data_type].append(load_spreadsheet(filename, sheet_name))
        else:
            file_data[data_type].append(False)
        print(file_data[data_type][2])
    ############################################################################

    

    ############################################################################
    # This section loops through each senior, each data_type, and each manager
    # 
    #
    ############################################################################
    for senior in senior_to_manager_dict.keys():  # loop through each senior manager
        print("\\\n-> Compiling data for {}".format(senior))  # print status
        senior_cumulative_gc_list = []  # total list of group codes for each senior manager; initialize as empty
        for data_type in file_data.keys():  # loop through each data_type to filter and work
            if file_data[data_type][2] is not False:  # meaning file exists in the folder
                for manager in senior_to_manager_dict[senior]:  # loop through each manager for the senior manager
                    print("--> Compiling {} data for {}".format(data_type, manager))  # print status

                    if "CG" in data_type:  # ETAC items
                        if manager in manager_to_gc_dict.keys():  # if manager group code data exists
                            senior_cumulative_gc_list.append(manager_to_gc_dict[manager]) # add that manager's group codes to the senior manager's group code list
                            manager_filtered_data = filter_etac_data(file_data[data_type][2], data_type, manager_to_gc_dict[manager])  # filter and return pandas DataFrame of all filtered data for the manager    
                        
                            if len(manager_filtered_data) > 0: 
                                add_to_email_body(manager_emailobj_dict[manager], data_type, manager_filtered_data)  #  create a table from the data and add it to the body of the email for each manager
                                add_to_email_body(senior_emailobj_dict[senior], data_type, manager_filtered_data, _name=manager)  #  create a table from the data and add it to the body of the email for each manager
    ###########################################################################
    
    

    ############################################################################
    # This section loops through each senior, and each manager, and sends 
    # each email via the email class method, then creates an .html for each page
    #
    ############################################################################
''' print("\n-> Sending emails and creating webpages")
    for senior in senior_to_manager_dict.keys():  # loop through each senior manager
        print("\n--> Executing for {}".format(senior))
        for manager in senior_to_manager_dict[senior]:  # loop through each manager for the senior manager
            print("---> Executing for {}".format(manager))
            if manager_emailobj_dict[manager].body != "":
                add_footer_to_body(manager_emailobj_dict[manager])
                #manager_emailobj_dict[manager].send_email()
            else:
                manager_emailobj_dict[manager].body = "<h1>Congratulations, you have either completed all tasks due within the next two weeks or you were not assigned tasks in this timeframe.</h1>"
            manager_emailobj_dict[manager].body = "<html><body>" + manager_emailobj_dict[manager].body + "</body></html>"
            fnc = open("\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\html_pages\\{}.html".format(manager), "w")
            fnc.write(manager_emailobj_dict[manager].body)
            fnc.close()
        if senior_emailobj_dict[senior].body != "":
            add_footer_to_body(senior_emailobj_dict[senior])
            #senior_emailobj_dict[senior].send_email()
        else:
            senior_emailobj_dict[senior].body = "<h1>Congratulations, you have either completed all tasks due within the next two weeks or you were not assigned tasks in this timeframe.</h1>"
        senior_emailobj_dict[senior].body = "<html><body>" + senior_emailobj_dict[senior].body + "</body></html>"
        fnc = open("\\\\nw\\data\\AVI\\RTB Weekly Metrics for Managers\\html_pages\\{}.html".format(senior), "w")
        fnc.write(senior_emailobj_dict[senior].body)
        fnc.close() '''
    ###########################################################################


print("\n\n\n-------------------SCRIPT COMPLETE!-------------------")
















"""
    ### Extracing ETAC Data
    for senior in manager_list.keys():
        first_lines = manager_list[senior]
        
        filtered_data = raw_data_etac[raw_data_etac["Manager"].str.contains(first_lines[0])]
        
        raw_data = filtered_data.reset_index(drop=True)
            
        parameters_ = []

        late_qty_due = 0
        this_week_qty_due = 0
        two_week_qty_due = 0

        late_qty_ecd = 0
        this_week_qty_ecd = 0
        two_week_qty_ecd = 0
        
        # this loop takes the numpy array and turns it into a dictionary to pass to jinja2 to render
        for i in range(len(raw_data)):

            if raw_data.OBDue[i] < datetime.today():
                late_qty_due = late_qty_due + 1
            elif (raw_data.OBDue[i] > datetime.today()) and (raw_data.OBDue[i] <= datetime.today() + timedelta(days=8)):
                this_week_qty_due = this_week_qty_due + 1
            if (raw_data.OBDue[i] > datetime.today()) and (raw_data.OBDue[i] <= datetime.today() + timedelta(days=14)):
                two_week_qty_due = two_week_qty_due + 1

            if raw_data.OBECD[i] < datetime.today():
                late_qty_ecd = late_qty_ecd + 1
            elif (raw_data.OBECD[i] > datetime.today()) and (raw_data.OBECD[i] <= datetime.today() + timedelta(days=8)):
                this_week_qty_ecd = this_week_qty_ecd + 1
            if (raw_data.OBECD[i] > datetime.today()) and (raw_data.OBECD[i] <= datetime.today() + timedelta(days=14)):
                two_week_qty_ecd = two_week_qty_ecd + 1  
            
            parameters_.append({'ID': raw_data.RelRecNo[i],
                                'Work_No': raw_data.WorkNo[i],
                                'Model': raw_data.Model[i],
                                'CG': str(raw_data.CG[i])[:2],
                                'Lead': raw_data.Lead[i],
                                'OBDue': str(raw_data.OBDue[i])[:10],
                                'OBECD': str(raw_data.OBECD[i])[:10],
                                'Description': raw_data.Description[i],
                                'Remarks': raw_data.Remarks[i],
                                'Manager': raw_data.Manager[i]})

        parameters_due = parameters_.copy()
        parameters_due.sort(key=operator.itemgetter('OBDue'))
        parameters_ecd = parameters_.copy()
        parameters_ecd.sort(key=operator.itemgetter('OBECD'))

        metrics = {'Late_qty_due': late_qty_due, 
                       'This_week_qty_due': this_week_qty_due,
                       'Two_week_qty_due': two_week_qty_due,
                       'Late_qty_ecd': late_qty_ecd, 
                       'This_week_qty_ecd': this_week_qty_ecd,
                       'Two_week_qty_ecd': two_week_qty_ecd,
                       'Updated_date': datetime.today(),
                       'Manager': senior}
    
        template_filename = "dwd_senior_template.html"
        rendered_filename = "dwd_{}.html".format(senior)

        template_file_dir = "C:\\Users\\kv542d\\Documents\\Desktop\\Management\\FTCM Website\\FTCM Website\\templates\\"
        rendered_file_dir = "C:\\Users\\kv542d\\Documents\\Desktop\\Management\\FTCM Website\\FTCM Website\\pages\\"

        render_environment = jinja2.Environment(loader=jinja2.FileSystemLoader(template_file_dir), trim_blocks=True)

        output_text = render_environment.get_template(template_filename).render(parameters_due=parameters_due, metrics=metrics,
                                                                                parameters_ecd=parameters_ecd)

        with open(rendered_file_dir+rendered_filename, "w", encoding='utf-8') as result_file:
            result_file.write(output_text)
    
    """