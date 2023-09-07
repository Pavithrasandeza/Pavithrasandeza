import warnings
warnings.simplefilter("ignore")
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

def percent(file_name,index_name,output_file,sla):
    pivot_sla=pd.pivot_table(file_name,index=['Assigned to'],columns=[sla],values=[],fill_value=0,observed=True,aggfunc=lambda x: np.count_nonzero(x == True))
    true_sla=file_name[sla].sum()
    total_sla=file_name[sla].count()
    sla_percent=int((true_sla/total_sla)*100)
    prev_total=int(input('enter the previous month total  for '+index_name+': '))
    prev_percent=int(input('enter the previous month sla percent for '+index_name+': '))
    df=pd.DataFrame()
    df[index_name] = ['Total', 'Sla']
    df.set_index(index_name,inplace = True)
    month_column={'JAN': ['DEC', 'JAN'],'FEB': ['JAN', 'FEB'],'MAR': ['FEB', 'MAR'],'APR': ['MAR', 'APR'],'MAY': ['APR', 'MAY'],'JUN': ['MAY', 'JUN'],'JUL': ['JUN', 'JUL'],'AUG': ['JUL', 'AUG'],'SEP': ['AUG', 'SEP'],'OCT': ['SEP', 'OCT'],'NOV': ['OCT', 'NOV'],'DEC': ['NOV', 'DEC']}
    for key,value in month_column.items():
        if key==month:
            column_name=value
            break
    df[column_name[0]] = None
    df[column_name[1]] = None
    df.loc['Total'] = [prev_total,total_sla]
    df.loc['Sla'] = [f'{prev_percent}%', f'{sla_percent}%']
    df.to_excel(output_file)
    if index_name=='Incidence':
        total_value=[prev_total,total_sla]
        percent_value=[prev_percent,sla_percent]
        fig, ax1 = plt.subplots()
        upper_limit=max(prev_total,total_sla)+200
        plt.bar(column_name,total_value,color ='blue') 
        tick_positions = np.arange(0, upper_limit+200,200)
        for x, y in enumerate(total_value):
            plt.text(x, y, str(y), ha='center', va='bottom')
        plt.yticks(tick_positions)
        plt.legend(labels=['Incidence'], loc='upper right', bbox_to_anchor=(0.5, -0.05),fancybox=True, shadow=True, ncol=2)

        lower_limit=min(prev_percent,sla_percent)-5
        upper_limit=max(prev_percent,sla_percent)+5
        ax2 = ax1.twinx()
        ax2.set_ylim(lower_limit,upper_limit)        
        tick_positions = np.arange(lower_limit, upper_limit + 5,5)
        plt.yticks(tick_positions)
        plt.plot(column_name,percent_value, color='red')
        for x, y in zip(column_name, percent_value):
            label = "{}%".format(y)
            plt.annotate(label,(x,y),textcoords="offset points",xytext=(0,10),ha='center')
        plt.title('SLA PROGRESS')
        plt.legend(labels=['SLA'],loc='upper left', bbox_to_anchor=(0.5, -0.05),fancybox=True, shadow=True, ncol=2)
        plt.show()

month=input("\nEnter the month for which DASHBOARD is created:")
incidence=pd.read_excel(r"C:\Users\Avita\Documents\sandeza\incidence.xlsx")
percent(incidence,'Incidence','Inci_percent.xlsx','Made SLA')
request=pd.read_excel(r"C:\Users\Avita\Documents\sandeza\sc_task.xlsx")
percent(request,'Request','Request_percent.xlsx','Active')
