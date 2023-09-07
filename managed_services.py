import warnings
warnings.simplefilter("ignore")
import numpy as np
import pandas as pd
workstream_group = {'Kodhandapani K': 'Windows','Hemalatha Venkatesan': 'Windows','David Jabez Devasingh S': 'Mac','Zakeer Ismail': 'Voice Security','Ganesh Kumar': 'Security','Prashanth Chandrasekar': 'Network','Shekhar Muddangula': 'Network','Suman Chetty': 'Network','Karthik Visvanathan': 'Servers-1','Tarani Teja': 'Servers-1','Satyanarayana Madiraju': 'Servers-2','Prabakaran Velayutham': 'Servers-2','Madiraju Kumar': 'Servers-3','Dilli Babu': 'Servers-3','Srivatsava Ganeshu': 'SSO'}

cr=pd.read_excel(r"C:\Users\Avita\Documents\May 2023\May 2023\change_request.xlsx")
pivot_cr=pd.pivot_table(cr,index=['Assigned to'],columns=['State'],values=[],fill_value=0,observed=True, aggfunc=np.count_nonzero)
pivot_cr['CRs']=pivot_cr['Canceled']+pivot_cr['Closed']
pivot_cr.drop(columns=['Canceled', 'Closed'], inplace=True)
pivot_cr.to_excel('cr.xlsx')

incidence=pd.read_excel(r"C:\Users\Avita\Documents\sandeza\incidence.xlsx")
pivot_sla=pd.pivot_table(incidence,index=['Assigned to'],columns=['Made SLA'],values=[],fill_value=0,observed=True,aggfunc=lambda x: np.count_nonzero(x == True))
pivot_sla.rename(columns={True: 'Incidents SLA Met'}, inplace=True)
pivot_sla.to_excel('sla.xlsx')

cr=pd.read_excel(r"cr.xlsx")
req=pd.read_excel(r"named_sctask.xlsx")
inci=pd.read_excel(r"named_incidence.xlsx")
sla=pd.read_excel(r"sla.xlsx")
service=inci[['Assigned to','Resolved','Pending']].merge(req[['Assigned to','Resolved','Pending']],on ='Assigned to', how = "left").merge(cr[['Assigned to','CRs']],on ='Assigned to',how = "left").merge(sla[['Assigned to','Incidents SLA Met']], on ='Assigned to',how = "left")
service.set_index('Assigned to',inplace = True)
service['Pending INC/RITM']=service['Pending_x']+service['Pending_y']
column_rename={'Resolved_x':'Incidents Resolved','Resolved_y':'Requests Resolved'}
service.rename(columns=column_rename,inplace=True)
service.drop(columns=['Pending_x', 'Pending_y'], inplace=True)
service.to_excel("managed_services.xlsx")

service.rename(index=workstream_group,inplace=True)
service=service.groupby(['Assigned to']).sum()
service.loc['Total']=service.sum(axis=0)
service.replace(0, '',inplace=True)
service.index.names = ['Work Stream']
service.to_excel('managed_services.xlsx')

spoc={'Work Stream':['Windows','Mac','Voice Security','Security','SSO','Network','Servers-1','Servers-3','Servers-2'],
      'Sandeza SPOC_L1':['Kodhandapani','David J','Zakeer','Ganesh','Srivatsava','Prashanth','Karthik','Madiraju','Satya'],
      'ServiceNow SPOC':['Dilnasheen','David B','Avinash','Ankit Shah','Ankit Shah','Avinash','Balaji','Balaji','Balaji']}
spoc_df = pd.DataFrame(spoc)
spoc_df.set_index('Work Stream',inplace = True)
counts = {}
for value in workstream_group.values():
    if value in counts:
        counts[value] += 1
    else:
        counts[value] = 1
size= pd.DataFrame(list(counts.items()), columns=['Work Stream', 'Sandeza TeamSize'])
size.set_index('Work Stream',inplace = True)
service=service.merge(spoc_df, left_index=True, right_index=True, how='left').merge(size, left_index=True, right_index=True, how='left')
service=service.assign(**{'RCAs': '','vulnerability': '','KBs Created': ''})
column_order=['Sandeza SPOC_L1','ServiceNow SPOC','Sandeza TeamSize','Incidents Resolved','Requests Resolved','CRs','RCAs','Incidents SLA Met','vulnerability','KBs Created','Pending INC/RITM']
service=service[column_order]
row_order=['Windows','Mac','SSO','Voice Security','Security','Servers-1','Servers-2','Servers-3','Network','Total']
service=service.loc[row_order]
service.to_excel('managed_services.xlsx')

