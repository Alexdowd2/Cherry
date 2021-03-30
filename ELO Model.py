#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import numpy as np
import sklearn
from sklearn.preprocessing import OneHotEncoder
ohe = OneHotEncoder(sparse = False)

df = pd.read_csv(r'filepath.csv')
df.columns


# In[ ]:


df = df[['date','season','team1','team2','elo1_pre',
       'elo2_pre','elo1_post', 'elo2_post',
       'qbelo1_pre', 'qbelo2_pre','qb1_value_pre',
       'qb2_value_pre','qb1_game_value', 'qb2_game_value', 'qb1_value_post', 'qb2_value_post',
       'qbelo1_post', 'qbelo2_post','score1','score2']]

def winning_team(row):
    if row['score1'] > row['score2']:
        global winner
        winner = 1
    elif row['score1'] < row['score2']:
        winner = 2
    return winner

df['winner'] = df.apply(winning_team, axis = 1)
df.tail()
df.fillna(0,inplace = True)
print(df.dtypes)


# In[ ]:


train = df[['elo1_pre',
       'elo2_pre','elo1_post', 'elo2_post',
       'qbelo1_pre', 'qbelo2_pre','qb1_value_pre',
       'qb2_value_pre','qb1_game_value', 'qb2_game_value', 'qb1_value_post', 'qb2_value_post',
       'qbelo1_post', 'qbelo2_post']]
train.astype('int64')


# In[ ]:


from sklearn import preprocessing
X = train
y = df['winner'].values


# In[ ]:


X= preprocessing.StandardScaler().fit(X).transform(X)
X[0:5]


# In[ ]:


from sklearn.model_selection import train_test_split

X_train, X_test, y_train, y_test = train_test_split( X, y, test_size=.2, random_state=4)
print ('Train set:', X_train.shape,  y_train.shape)
print ('Test set:', X_test.shape,  y_test.shape)


# In[ ]:


from sklearn.linear_model import LogisticRegression
LR = LogisticRegression(C=0.01, solver='liblinear').fit(X_train,y_train)
LR


# In[ ]:


yhat = LR.predict(X)
yhat


# In[ ]:


df_pred = pd.DataFrame(yhat)


# In[ ]:


df_pred['predicted winner'] = df_pred

df = pd.merge(df, df_pred['predicted winner'], how = 'left', left_index = True, right_index = True)


# In[ ]:


df.drop(['winner'], axis = 1, inplace = True)
df.tail()


# In[ ]:


df.to_csv('filepath', index = False)


# In[ ]:




