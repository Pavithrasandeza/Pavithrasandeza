import warnings
warnings.simplefilter("ignore")
import openpyxl
import numpy as np
import pandas as pd
closed_inci=pd.read_excel(r"C:\Users\Avita\Documents\May 2023\May 2023\Closed Incidents.xlsx")
pending_inci=pd.read_excel(r"C:\Users\Avita\Documents\May 2023\May 2023\Pending Incidents.xlsx")
merged_inci=pd.concat([closed_inci,pending_inci], ignore_index=True, sort=False)
merged_inci.to_excel('incidence.xlsx')     
closed_inci.head()
incidence=pd.read_excel(r"incidence.xlsx")

pivot_inci=pd.pivot_table(incidence,index=['Assigned to'],columns=['State'],values=[],fill_value=0,observed=True, aggfunc=np.count_nonzero)
pivot_inci['Resolved']=pivot_inci['Resolved']+pivot_inci['Cancelled']+pivot_inci['Closed']
pivot_inci['Awaiting Caller Info']=pivot_inci['Awaiting Caller Info']+pivot_inci['Awaiting Equipment']
column_rename={'New':'Open','Pending Vendor':'Pending','Work in Progress':'WIP'}
pivot_inci.rename(columns=column_rename,inplace=True)
pivot_inci.drop(columns=['Awaiting Equipment', 'Cancelled', 'Closed'], inplace=True)
pivot_inci['Total']=pivot_inci.sum(axis=1)
pivot_inci.to_excel('named_incidence.xlsx')

workstream_group = {'Kodhandapani K': 'Windows','Hemalatha Venkatesan': 'Windows','David Jabez Devasingh S': 'Mac','Zakeer Ismail': 'Voice Security','Ganesh Kumar': 'Security','Prashanth Chandrasekar': 'Network','Shekhar Muddangula': 'Network','Suman Chetty': 'Network','Karthik Visvanathan': 'Servers-1','Tarani Teja': 'Servers-1','Satyanarayana Madiraju': 'Servers-2','Prabakaran Velayutham': 'Servers-2','Madiraju Kumar': 'Servers-3','Dilli Babu': 'Servers-3','Srivatsava Ganeshu': 'SSO'}
pivot_inci.rename(index=workstream_group, inplace=True)
index_rename={'Windows':'Windows & MAC Endpoints','Mac':'Windows & MAC Endpoints','Servers-1':'Windows & Linux Server','Servers-2':'Windows & Linux Server','Servers-3':'Windows & Linux Server'}
pivot_inci.rename(index=index_rename,inplace=True)
pivot_inci=pivot_inci.groupby(['Assigned to']).sum()
pivot_inci.loc['Total']=pivot_inci.sum(axis=0)
pivot_inci.index.name = None
pivot_inci.to_excel('overall_incidence.xlsx')

import matplotlib.pyplot as plt
a=pivot_inci.T.plot.barh(stacked=True,title='Overall incidents')
a.xaxis.set_visible(False)
a.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05),fancybox=True, shadow=True, ncol=7)
pivot_inci=pivot_inci.replace(0, '')
pivot_inci.to_excel('overall_incidence.xlsx')
plt.show()

x=pivot_inci.loc['Total']
df = pd.DataFrame(x)
ax=df.plot(kind='barh', legend=False,title='Overall incidents')
for index, value in enumerate(x):
    ax.text(value, index, str(value))
plt.show()