import warnings
warnings.simplefilter("ignore")

import pandas as pd
closed_sc=pd.read_excel(r"C:\Users\Avita\Documents\May 2023\May 2023\SC Task Closed.xlsx")
pending_sc=pd.read_excel(r"C:\Users\Avita\Documents\May 2023\May 2023\SC Task Pending.xlsx")
merged_inci=pd.concat([closed_sc,pending_sc], ignore_index=True, sort=False)
merged_inci.to_excel('sc_task.xlsx')     
closed_sc.head()
task=pd.read_excel(r"sc_task.xlsx")

import numpy as np
task_pivot=pd.pivot_table(task,index=['Assigned to'],columns=['State'],values=[],fill_value=0,observed=True, aggfunc=np.count_nonzero)
task_pivot['Resolved']=task_pivot['Closed Complete']+task_pivot['Closed Incomplete']+task_pivot['Closed Skipped']
task_pivot.rename(columns={'Work in Progress':'WIP'},inplace=True)
task_pivot.drop(columns=['Closed Complete', 'Closed Incomplete', 'Closed Skipped'], inplace=True)
task_pivot['Total']=task_pivot.sum(axis=1)
task_pivot.to_excel('named_sctask.xlsx')

workstream_group = {'Kodhandapani K': 'Windows','Hemalatha Venkatesan': 'Windows','David Jabez Devasingh S': 'Mac','Zakeer Ismail': 'Voice Security','Ganesh Kumar': 'Security','Prashanth Chandrasekar': 'Network','Shekhar Muddangula': 'Network','Suman Chetty': 'Network','Karthik Visvanathan': 'Servers-1','Tarani Teja': 'Servers-1','Satyanarayana Madiraju': 'Servers-2','Prabakaran Velayutham': 'Servers-2','Madiraju Kumar': 'Servers-3','Dilli Babu': 'Servers-3','Srivatsava Ganeshu': 'SSO'}
task_pivot.rename(index=workstream_group, inplace=True)
index_rename={'Windows':'Windows & MAC Endpoints','Mac':'Windows & MAC Endpoints','Servers-1':'Windows & Linux Server','Servers-2':'Windows & Linux Server','Servers-3':'Windows & Linux Server'}
task_pivot.rename(index=index_rename,inplace=True)
task_pivot=task_pivot.groupby(['Assigned to']).sum()
task_pivot.loc['Total']=task_pivot.sum(axis=0)
task_pivot.index.name = None
task_pivot.to_excel('overall_requests.xlsx')

import matplotlib.pyplot as plt
a=task_pivot.T.plot.barh(stacked=True,title='Overall Requests')
a.xaxis.set_visible(False)
a.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05),fancybox=True, shadow=True, ncol=6)
task_pivot=task_pivot.replace(0, '')
task_pivot.to_excel('overall_requests.xlsx')
plt.show()

x=task_pivot.loc['Total']
df = pd.DataFrame(x)
ax=df.plot(kind='barh', legend=False,title='Overall Requests')
for index, value in enumerate(x):
    ax.text(value, index, str(value))
plt.show()