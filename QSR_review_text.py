import pandas as pd
import os
import re
import nltk
from nltk.corpus import stopwords
from nltk import word_tokenize
from nltk import pos_tag
from textblob import TextBlob
import warnings
warnings.filterwarnings("ignore")
from collections import Counter
from operator import itemgetter
import itertools

#clean the text of any special or alpha-numeric characters
def clean_text(review_text):
    review_text = re.sub('[^A-Za-z0-9]+', ' ', review_text)
    return review_text

def count_pos(words):
    count_pos = {}
    for k in range(0, len(words)):
        split_pos = []
        split_pos = nltk.pos_tag(words[k].split())
        if count_pos == {}:
            count_pos.update(dict(Counter(elem[0] for elem in split_pos if 
    			elem[1].startswith('J') or elem[1].startswith('R') or elem[1].startswith('V'))))
        else: 
            count_k = dict(Counter(elem[0] for elem in split_pos if 
    elem[1].startswith('J') or elem[1].startswith('R') or elem[1].startswith('V')))
            count_pos = {k: count_k.get(k, 0) + count_pos.get(k, 0) for k in set(count_k) | set(count_pos)}
    count_pos = dict(sorted(count_pos.items(), key = itemgetter(1), reverse = True))
    return count_pos

#get file location
folder = input('Enter folder path: ')
file = input('Enter file name: ')
file = file + '.xlsx'
sheet = input('Enter sheet name: ')
df = pd.read_excel(os.path.join(folder, file), sheet_name = sheet) #create dataframe
qsr_df = df.copy()
qsr_df.dropna(subset=['review_text'], inplace=True) #remove nan
qsr_df = qsr_df.reset_index(drop=True) #reindex from 0
qsr_df['day'] = qsr_df['date'].dt.day #get day from date
#clean review text
qsr_df['clean_review_text'] = qsr_df['review_text'].apply(clean_text) #remove alphanumeric and special characters
qsr_df['clean_review_text'] = qsr_df['clean_review_text'].str.lower() #all lower case
#stop words but include certain necessary words
stop = stopwords.words('english')
stop.remove('not')
stop.remove('no')
stop.remove("don't")
stop.remove("doesn't")
stop.append('pizza')
stop.append('biryani')
stop.append('chicken')
stop.append('burger')
stop.append('burgers')
stop.append('momo')
#remove stop words in clean_review_text column
qsr_df['clean_review_text'] = qsr_df['clean_review_text'].apply(lambda x: " ".join(x for x in x.split() if x not in stop))
qsr_df['res_name'] = qsr_df.res_name.str.strip()
#spell correction
qsr_df['clean_review_text'] = qsr_df['clean_review_text'].apply(lambda x: ''.join(TextBlob(x).correct()))
#tokenize clean_review_text
qsr_df['crt_tokenized'] = qsr_df['clean_review_text'].apply(word_tokenize)
#apply part of speech tag
qsr_df['crt_pos'] = qsr_df['crt_tokenized'].apply(pos_tag)
#get polarity
qsr_df['polarity'] = qsr_df['clean_review_text'].apply(lambda x: TextBlob(x).sentiment[0])
#negative, positive or neutral for each review
qsr_df['analysis'] = 0 * len(qsr_df)
for i in range(0, len(qsr_df)):
    if qsr_df['polarity'][i] <= -0.5:
        qsr_df['analysis'][i] = 0
    elif qsr_df['polarity'][i] > -0.5 and qsr_df['polarity'][i] <= -0.25:
        qsr_df['analysis'][i] = 1
    elif qsr_df['polarity'][i] > -0.25 and qsr_df['polarity'][i] < 0:
        qsr_df['analysis'][i] = 2
    elif qsr_df['polarity'][i] == 0:
        qsr_df['analysis'][i] = 3
    elif qsr_df['polarity'][i] > 0 and qsr_df['polarity'][i] <= 0.5:
        qsr_df['analysis'][i] = 4
    else: 
        qsr_df['analysis'][i] = 5
        
qsr_df['city_res_name'] = qsr_df['res_name'] + '-' + qsr_df['city_name'] #concatenate city and restaurant name
polarity_rating = qsr_df.groupby(['rating']).mean()['polarity'] #polarity for each rating
#review for each comment
qsr_df['review'] = ' '
for i in range(0, len(qsr_df)):
    if qsr_df['polarity'][i] <= 0.01 or qsr_df['rating'][i] < 3:
        qsr_df['review'][i] = 'Negative'
    else:
        qsr_df['review'][i] = 'Positive'
#polarity for each day
polarity_day = qsr_df.groupby(['day', 'review']).count().reset_index()
polarity_day = polarity_day[['day', 'review', 'date']]
polarity_day_2 = (polarity_day['review'] == 'Negative')
polarity_day_3 = polarity_day[polarity_day_2]
polarity_day_3.rename(columns={'date':'number_complaints'}, inplace = True)
polarity_day_3 = polarity_day_3.sort_values('number_complaints', ascending = False) #final polarity by day
polarity_day_3 = polarity_day_3.reset_index(drop = True)
#max negative and max positive complaints by restaurant and city
city_res_pol = qsr_df.groupby(['res_name', 'city_name', 'review']).count().reset_index()

city_res_pol_neg = city_res_pol[['city_name', 'res_name', 'review', 'date']]
city_res_pol_neg = (city_res_pol['review'] == 'Negative')
city_res_pol_neg_2 = city_res_pol[city_res_pol_neg]
city_res_pol_neg_2.rename(columns={'date':'number_complaints'}, inplace = True)
city_res_pol_neg_2 = city_res_pol_neg_2.sort_values('number_complaints', ascending = False) #final polarity
city_res_pol_neg_2 = city_res_pol_neg_2[['res_name', 'city_name', 'number_complaints']]
city_res_pol_neg_2 = city_res_pol_neg_2.reset_index(drop = True)

city_res_pol_pos = city_res_pol[['city_name', 'res_name', 'review', 'date']]
city_res_pol_pos = (city_res_pol['review'] == 'Positive')
city_res_pol_pos_2 = city_res_pol[city_res_pol_pos]
city_res_pol_pos_2.rename(columns={'date':'number_compliments'}, inplace = True)
city_res_pol_pos_2 = city_res_pol_pos_2.sort_values('number_compliments', ascending = False) #final polarity
city_res_pol_pos_2 = city_res_pol_pos_2[['res_name', 'city_name', 'number_compliments']]
city_res_pol_pos_2 = city_res_pol_pos_2.reset_index(drop = True)
#cities with max positive and negative comments
max_pos_neg = qsr_df.groupby('review')['city_name'].apply(lambda x: x.value_counts())
#negative and positive words and counts associated with each restaurant/brand
res_group = pd.DataFrame(qsr_df.groupby('res_name'))
res_group['neg_words'] = ' '
res_group['pos_words'] = ' '
res_group['top5_neg_words'] = ' '
res_group['top5_pos_words'] = ' '
for i in range(0, len(res_group)):
    pos_words = []
    neg_words = []
    res_group[1][i] = res_group[1][i].reset_index(drop = True)
    for j in range(0, len(res_group[1][i])):
        if res_group[1][i]['polarity'][j] <= 0.01 or res_group[1][i]['rating'][j] <= 3:
            neg_words.append(res_group[1][i]['clean_review_text'][j])
        else:
            pos_words.append(res_group[1][i]['clean_review_text'][j])
    if neg_words == []:
        res_group['neg_words'][i] = ' '
        res_group['top5_neg_words'][i] = ' '
    elif neg_words != []:
        res_group['neg_words'][i] = count_pos(neg_words)
        if len(res_group['neg_words'][i]) < 5:
            res_group['top5_neg_words'][i] = res_group['neg_words'][i]
        elif len(res_group['neg_words'][i]) >= 5:
            res_group['top5_neg_words'][i] = dict(itertools.islice(res_group['neg_words'][i].items(), 5))
    if pos_words == []:
        res_group['pos_words'][i] = ' '
        res_group['top5_pos_words'][i] = ' '
    elif pos_words != []:
        res_group['pos_words'][i] = count_pos(pos_words)
        if len(res_group['pos_words'][i]) < 5:
            res_group['top5_pos_words'][i] = res_group['pos_words'][i]
        elif len(res_group['pos_words'][i]) >= 5:
            res_group['top5_pos_words'][i] = dict(itertools.islice(res_group['pos_words'][i].items(), 5))

res_group = res_group.rename(columns = {0: 'res_name'})
res_group_neg = res_group[['res_name', 'top5_neg_words']]
res_group_pos = res_group[['res_name', 'top5_pos_words']]
#write output to Excel file
out_file = file[:-5] + ' Output.xlsx'
with pd.ExcelWriter(out_file) as writer:
    polarity_day_3.to_excel(writer, sheet_name = 'days_max_complaints') #q1
    city_res_pol_neg_2.to_excel(writer, sheet_name = 'rest_max_negative') #q2
    city_res_pol_pos_2.to_excel(writer, sheet_name = 'rest_max_positive') #q2
    res_group_neg.to_excel(writer, sheet_name = 'brand_negative') #q3
    res_group_pos.to_excel(writer, sheet_name = 'brand_positive') #q4
    max_pos_neg.to_excel(writer, sheet_name = 'city_max_positive _negative') #q5