import pandas as pd
import os
import numpy as np
import warnings
warnings.filterwarnings("ignore")

#get file location
folder = input('Enter folder path: ')
file = input('Enter file name: ')
file = file + '.csv'
df = pd.read_csv(os.path.join(folder, file)) #create dataframe
ssb_df = df.copy()
#cleaning dataframe -> update headers, fill nan with 0, drop empty rows
ssb_df = ssb_df.dropna(axis = 0, how = 'all')
ssb_df = ssb_df.reset_index(drop = True)
ssb_df = ssb_df.T
ssb_df.reset_index(drop = True, inplace = True)
ssb_df = ssb_df.fillna(0)
new_header = ssb_df.iloc[0]
ssb_df = ssb_df[1:]
ssb_df.columns = new_header
#q1 - %gross and net sales of each brand
gross_net_sales = ssb_df.iloc[:, np.r_[0, 1, 4]]
gross_net_sales['perc_gross_sales'] = pd.to_numeric(gross_net_sales.iloc[:,1]) / float(gross_net_sales.iloc[:,1][1]) * 100
gross_net_sales['perc_net_sales'] = pd.to_numeric(gross_net_sales.iloc[:,2]) / float(gross_net_sales.iloc[:,2][1]) * 100
gross_sales = gross_net_sales.iloc[:, np.r_[0, 1, 3]]
gross_sales = gross_sales.sort_values(['perc_gross_sales'], ascending = [False])
gross_sales = gross_sales.reset_index(drop = True)
net_sales = gross_net_sales.iloc[:, np.r_[0, 1, 4]]
net_sales = net_sales.sort_values(['perc_net_sales'], ascending = [False])
net_sales = net_sales.reset_index(drop = True)
#q2 - brand wise discount %
discount_perc = ssb_df.iloc[:,np.r_[0, 3]]
discount_perc['perc_brand_disc'] = pd.to_numeric(discount_perc.iloc[:,1]) / float(discount_perc.iloc[:,1][1]) * 100
discount_perc = discount_perc.iloc[:, np.r_[0, 2]]
discount_perc = discount_perc.sort_values(['perc_brand_disc'], ascending = [False])
discount_perc = discount_perc.reset_index(drop = True)
#q3 - packing charge % by brand
pack_charge = ssb_df.iloc[:, np.r_[0, 19]]
pack_charge['pack_charge'] = pd.to_numeric(pack_charge.iloc[:,1]) / float(pack_charge.iloc[:,1][1]) * 100
pack_charge = pack_charge.iloc[:, np.r_[0, 2]]
pack_charge = pack_charge.sort_values(['pack_charge'], ascending = [False])
pack_charge = pack_charge.reset_index(drop = True)
#q4 - brand wise contribution to total transactions and % of total transactions
brand_transact = ssb_df.iloc[:, np.r_[0, 24, 25]]
brand_transact['total_transactions'] = pd.to_numeric(brand_transact.iloc[:,1]) * pd.to_numeric(brand_transact.iloc[:,2])
brand_transact['number_transact_perc'] = pd.to_numeric(brand_transact.iloc[:,1]) / float(brand_transact.iloc[:,1][1]) * 100
brand_transact['total_transact_perc'] = pd.to_numeric(brand_transact.iloc[:,3]) / float(brand_transact.iloc[:,3][1]) * 100
number_transact = brand_transact.iloc[:, np.r_[0, 4]]
number_transact = number_transact.sort_values(['number_transact_perc'], ascending = False)
number_transact = number_transact.reset_index(drop = True)
amount_transact = brand_transact.iloc[:, np.r_[0, 5]]
amount_transact = amount_transact.sort_values(['total_transact_perc'], ascending = False)
amount_transact = amount_transact.reset_index(drop = True)
avg_transact = brand_transact.iloc[:,np.r_[0, 2]]
avg_transact[' Avg Sale Per Transaction'] = pd.to_numeric(avg_transact[' Avg Sale Per Transaction'])
avg_transact = avg_transact.sort_values([' Avg Sale Per Transaction'], ascending = False)
avg_transact.loc[len(avg_transact)] = ['Average', brand_transact.iloc[0,2]]
avg_transact = avg_transact.reset_index(drop = True)
#q5 - acceptance rate deviation from 100%
accept_rate = ssb_df.iloc[:, np.r_[0, 27]]
accept_rate['deviation'] = ((pd.to_numeric(accept_rate.iloc[:, 1]) - 100) / 100) * 100
accept_rate = accept_rate.iloc[1:, :]
accept_rate = accept_rate.sort_values(['deviation'], ascending = False)
accept_rate = accept_rate.reset_index(drop = True)
#q6 - process rate deivation from 100%
process_rate = ssb_df.iloc[:, np.r_[0, 28]]
process_rate['deviation'] = ((pd.to_numeric(process_rate.iloc[:,1]) - 100) / 100) * 100
process_rate = process_rate.iloc[1:, :]
process_rate = process_rate.sort_values('deviation', ascending = False)
process_rate = process_rate.reset_index(drop = True)
#q7 - acceptance, process, dispatch and delivery time deviation from average
brand_time = ssb_df.iloc[:, np.r_[0, 29, 30, 31, 32]]
avg_acc_time = pd.to_numeric(brand_time['Time To Accept (Mins)']).mean()
avg_process_time = pd.to_numeric(brand_time['Time To Process (Mins)']).mean()
avg_dispatch_time = pd.to_numeric(brand_time['Time To Dispatch (Mins)']).mean()
avg_deliver_time = pd.to_numeric(brand_time['Time To Deliver (Mins)']).mean()
brand_time['acc_time_dev'] = (pd.to_numeric(brand_time['Time To Accept (Mins)']) - avg_acc_time) / avg_acc_time * 100
brand_time['process_time_dev']  = (pd.to_numeric(brand_time['Time To Process (Mins)']) - avg_process_time) / avg_process_time * 100
brand_time['dispatch_time_dev'] = (pd.to_numeric(brand_time['Time To Dispatch (Mins)']) - avg_dispatch_time) / avg_dispatch_time * 100
brand_time['deliver_time_dev'] = (pd.to_numeric(brand_time['Time To Deliver (Mins)']) - avg_deliver_time) / avg_deliver_time * 100
brand_time = brand_time.iloc[1:, :]
acc_dev = brand_time.iloc[:, np.r_[0, 1, 5]]
acc_dev = acc_dev.sort_values('acc_time_dev', ascending = False)
acc_dev.loc[-1] = ['Average', avg_acc_time, 0]
acc_dev = acc_dev.reset_index(drop = True)
process_dev = brand_time.iloc[:, np.r_[0, 2, 6]]
process_dev = process_dev.sort_values('process_time_dev', ascending = False)
process_dev.loc[-1] = ['Average', avg_process_time, 0]
process_dev = process_dev.reset_index(drop = True)
dispatch_dev = brand_time.iloc[:, np.r_[0, 3, 7]]
dispatch_dev = dispatch_dev.sort_values('dispatch_time_dev', ascending = False)
dispatch_dev.loc[-1] = ['Average', avg_dispatch_time, 0]
dispatch_dev = dispatch_dev.reset_index(drop = True)
deliver_dev = brand_time.iloc[:, np.r_[0, 4, 8]]
deliver_dev = deliver_dev.sort_values('deliver_time_dev', ascending = False)
deliver_dev.loc[-1] = ['Average', avg_deliver_time, 0]
deliver_dev = deliver_dev.reset_index(drop = True)
#q8 - top 20 products
top_prod = ssb_df.iloc[:, np.r_[0, 40:128]]
top_prod = top_prod.T
top_prod = top_prod.iloc[:, 0]
top_prod = pd.DataFrame(top_prod)
top_prod.reset_index(level = 0, inplace = True)
new_head = top_prod.iloc[0]
top_prod = top_prod[1:]
top_prod.columns = new_head
top_prod.iloc[:,1] = pd.to_numeric((top_prod.iloc[:,1]))
top_prod = top_prod.nlargest(21, ['Total'])
top_prod['perc_sales'] = top_prod['Total'] / top_prod['Total'][1] * 100
top_prod = top_prod.iloc[1:, :]
top_prod = top_prod.reset_index(drop = True)
#output data
out_file = file[:-4] + ' Output.xlsx'
with pd.ExcelWriter(out_file) as writer:
    gross_sales.to_excel(writer, sheet_name = '% total gross sales') #q1
    net_sales.to_excel(writer, sheet_name = '% total net sales') #q1
    discount_perc.to_excel(writer, sheet_name = '% brand wise discount') #q2
    pack_charge.to_excel(writer, sheet_name = '% brand pack charge') #q3
    number_transact.to_excel(writer, sheet_name = '% brand no. transactions') #q4
    amount_transact.to_excel(writer, sheet_name = '% brand amount transactions') #q4
    avg_transact.to_excel(writer, sheet_name = 'brand avg sale per transaction') #q4
    accept_rate.to_excel(writer, sheet_name = 'acc rate % dev from 100%') #q5
    process_rate.to_excel(writer, sheet_name = 'process rate % dev from 100%') #q6
    acc_dev.to_excel(writer, sheet_name = 'acc rate % from avg acc time') #q7
    process_dev.to_excel(writer, sheet_name = 'pro rate % from avg pro time') #q7
    dispatch_dev.to_excel(writer, sheet_name = 'dis rate % from avg dis time') #q7
    deliver_dev.to_excel(writer, sheet_name = 'del rate % from avg del time') #q7
    top_prod.to_excel(writer, sheet_name = 'top 20 products by gross sales') #q8