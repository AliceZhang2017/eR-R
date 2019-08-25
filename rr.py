
# coding: utf-8

import pandas as pd
import glob
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import glob

os.chdir('/Users/yahuizhang/Desktop/python/AMJ')
FileList = glob.glob('*.xlsx')
# Initialize empty dataframe
df = pd.DataFrame()
# Loop over list of Excel files, import into dataframe, add date field, and export
for f in FileList:
    df = pd.read_excel(f, skiprows=0, skipfooter=0)
    all_data = all_data.append(df,ignore_index=True)

#Calalute P&G brand postive rate on a category level for both product and non-product
category_coment = all_data.groupby(['category','brand', 'is_competitor', 'emotion'],                                    as_index=False)['comments#'].sum()
category = pd.DataFrame(category_coment)
# pg positive% brand level
pg = pd.DataFrame(category[category.is_competitor == 0])
pg_pivot = pg.pivot_table(values = 'comments#', index = ['category', 'brand'], columns= ['emotion'])
pg_pivot['Total'] = pg_pivot['Negative'] +  pg_pivot['Positive'] +  pg_pivot['Neutral']  
pg_pivot['Positive%'] =  pg_pivot['Positive'] / pg_pivot['Total'] 
pg_pivot['Negative%'] =  pg_pivot['Negative'] / pg_pivot['Total'] 
pg_pivot = pg_pivot.reset_index()

# pg positive% category level
pgc = pd.DataFrame(category[category.is_competitor == 0])
pgc_pivot = pgc.pivot_table(values = 'comments#', index = ['category'], columns= ['emotion'])
pgc_pivot['Total'] = pgc_pivot['Negative'] +  pgc_pivot['Positive'] +  pgc_pivot['Neutral']  
pgc_pivot['Positive%'] =  pgc_pivot['Positive'] / pgc_pivot['Total'] 
pgc_pivot['Negative%'] =  pgc_pivot['Negative'] / pgc_pivot['Total'] 

#category positive
ca = pd.DataFrame(category)
ca_pivot = ca.pivot_table(values = 'comments#', index = ['category'], columns= ['emotion'])
ca_pivot['Total'] = ca_pivot['Negative'] +  ca_pivot['Positive'] +  ca_pivot['Neutral']  
ca_pivot['Positive%'] =  ca_pivot['Positive'] / ca_pivot['Total'] 
ca_pivot['Negative%'] =  ca_pivot['Negative'] / ca_pivot['Total'] 

# comp positive%
co = pd.DataFrame(category[category.is_competitor == 1])
co_pivot = co.pivot_table(values = 'comments#', index = ['category'], columns= ['emotion'])
co_pivot['Total'] = co_pivot['Negative'] +  co_pivot['Positive'] +  co_pivot['Neutral']  
co_pivot['Positive%'] =  co_pivot['Positive'] / co_pivot['Total'] 
co_pivot['Negative%'] =  co_pivot['Negative'] / co_pivot['Total'] 

#reset index and calculate benchmark
pgc1 = pgc_pivot.reset_index()
co1 = co_pivot.reset_index()
ca1 = ca_pivot.reset_index()

#merge all pg, competitor, and category 
combine = pd.merge(pgc1, co1, right_index=True, left_index=True)
combine1 = pd.merge(combine, ca1, right_index=True, left_index=True)
combine2 =combine1 [ ['category_x', 'Positive%_x', 'Negative%_x','Positive%_y', 'Negative%_y', 'Positive%', 'Negative%' ]]
combine2.columns = ['Category', 'P&G_Pos%', 'P&G_Neg%','Com_Pos%', 'Com_Neg%', 'Indus_Pos%', 'Indus_Neg%']
combine2['Benchmark_Pos'] = combine2[['Com_Pos%', 'Indus_Pos%']].max(axis = 1)
combine2['Benchmark_Neg'] = combine2[['Com_Neg%', 'Indus_Neg%']].min(axis = 1)

final_all = pd.merge(pg_pivot, combine2, left_on = 'category', right_on = 'Category',     \
                     how = 'outer')
Result_all = final_all[['category', 'brand', 'P&G_Pos%', 'Benchmark_Pos','P&G_Neg%',      \
                        'Benchmark_Neg']]
Result_all['vs_Benchmark_Pos'] = Result['P&G_Pos%'] - Result['Benchmark_Pos']
Result_all['vs_Benchmark_Neg'] = Result['P&G_Neg%'] - Result['Benchmark_Neg']


#Calalute P&G brand postive rate on a category level for both product and package and function only
category_coment1 = pd.DataFrame(all_data.groupby(['category','brand', 'is_competitor', 'emotion', 'review_type'],                                    as_index=False)['comments#'].sum())
#subset review type in product package function
category1 = category_coment1[(category_coment1['review_type'] == 'Product')  \
                             | (category_coment1['review_type'] == 'Package')  \
                             |(category_coment1['review_type'] == 'Function')]
category2 = category1.groupby(['category','brand', 'is_competitor', 'emotion'],  \
                              as_index=False)['comments#'].sum()

# pg positive% brand level
pg = pd.DataFrame(category2[category2.is_competitor == 0])
pg_pivot = pg.pivot_table(values = 'comments#', index = ['category', 'brand'], columns= ['emotion'])
pg_pivot['Total'] = pg_pivot['Negative'] +  pg_pivot['Positive'] +  pg_pivot['Neutral']  
pg_pivot['Positive%'] =  pg_pivot['Positive'] / pg_pivot['Total'] 
pg_pivot['Negative%'] =  pg_pivot['Negative'] / pg_pivot['Total'] 
pg_pivot_p = pg_pivot.reset_index()

# pg positive% category level
pgc = pd.DataFrame(category2[category2.is_competitor == 0])
pgc_pivot = pgc.pivot_table(values = 'comments#', index = ['category'], columns= ['emotion'])
pgc_pivot['Total'] = pgc_pivot['Negative'] +  pgc_pivot['Positive'] +  pgc_pivot['Neutral']  
pgc_pivot['Positive%'] =  pgc_pivot['Positive'] / pgc_pivot['Total'] 
pgc_pivot['Negative%'] =  pgc_pivot['Negative'] / pgc_pivot['Total'] 

#category positive
ca = pd.DataFrame(category2)
ca_pivot = ca.pivot_table(values = 'comments#', index = ['category'], columns= ['emotion'])
ca_pivot['Total'] = ca_pivot['Negative'] +  ca_pivot['Positive'] +  ca_pivot['Neutral']  
ca_pivot['Positive%'] =  ca_pivot['Positive'] / ca_pivot['Total'] 
ca_pivot['Negative%'] =  ca_pivot['Negative'] / ca_pivot['Total'] 

# competitor positive%
co = pd.DataFrame(category2[category2.is_competitor == 1])
co_pivot = co.pivot_table(values = 'comments#', index = ['category'], columns= ['emotion'])
co_pivot['Total'] = co_pivot['Negative'] +  co_pivot['Positive'] +  co_pivot['Neutral']  
co_pivot['Positive%'] =  co_pivot['Positive'] / co_pivot['Total'] 
co_pivot['Negative%'] =  co_pivot['Negative'] / co_pivot['Total'] 

#reset index and calculate benchmark. benchmark max of competitor pos% and category pos%.  negative benchmark is min of 
#competitor neg% and category neg%
pgc1 = pgc_pivot.reset_index()
co1 = co_pivot.reset_index()
ca1 = ca_pivot.reset_index()
combine = pd.merge(pgc1, co1, right_index=True, left_index=True)
combine1 = pd.merge(combine, ca1, right_index=True, left_index=True)
combine_p =combine1 [ ['category_x', 'Positive%_x', 'Negative%_x','Positive%_y', 'Negative%_y', 'Positive%', 'Negative%' ]]
combine_p.columns = ['Category', 'P&G_Pos%', 'P&G_Neg%','Com_Pos%', 'Com_Neg%', 'Indus_Pos%', 'Indus_Neg%']
combine_p['Benchmark_Pos'] = combine2[['Com_Pos%', 'Indus_Pos%']].max(axis = 1)
combine_p['Benchmark_Neg'] = combine2[['Com_Neg%', 'Indus_Neg%']].min(axis = 1)

final_p = pd.merge(pg_pivot_p, combine_p, left_on = 'category', right_on = 'Category', how = 'outer')
Result = final_p[['category', 'brand', 'P&G_Pos%', 'Benchmark_Pos','P&G_Neg%',                   'Benchmark_Neg' ]]
Result['vs_Benchmark_Pos'] = Result['P&G_Pos%'] - Result['Benchmark_Pos']
Result['vs_Benchmark_Neg'] = Result['P&G_Neg%'] - Result['Benchmark_Neg']

##Clean up the format. As I only need a few columns from result product and result all 

Result_Product = Result[['category', 'brand', 'P&G_Pos%', 'vs_Benchmark_Pos', 'P&G_Neg%',
                         'vs_Benchmark_Neg']]
Result_all = Result_all[['category', 'brand', 'P&G_Pos%', 'vs_Benchmark_Pos', 'P&G_Neg%',
                          'vs_Benchmark_Neg']]
Output = pd.merge(Result_all, Result_Product, left_on = ['category','brand'], right_on =                   ['category','brand'],
                  how = 'outer'
                 )
##Subsetting needed columns
Output = Output[['category', 'brand', 'P&G_Pos%_x', 'vs_Benchmark_Pos_x',
                 'P&G_Pos%_y', 'vs_Benchmark_Pos_y', 'P&G_Neg%_y', 'vs_Benchmark_Neg_y'
                ]]
##Rename the columns
Output.columns = ['Category','Brand', 'TTL Pos%', 'vs_Max(Com,Cate)', 'Product_Pos%',                  'vs_Max(Com,Cate)','Product_Neg%', 'vs_Min(Com,Cate)' ]

Output.to_csv ('/Users/yahuizhang/Desktop/python/AMJ/Output.csv', index = None, header=True) 

