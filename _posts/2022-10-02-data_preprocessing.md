## Import necessary packages


```python
import pandas as pd
import xlwings as xw
```

## Loading the data
- using xlwings package
- manipulate excel data (read & write) using xlwings packages
- for more information refer to the official document `docs.xlwings.org/`


```python
path = "./rawdata/202209_월간시계열.xlsx"
wb = xw.Book(path)
sheet = wb.sheets['1.매매종합']
```


```python
row_num = sheet.range(1,1).end('down').end('down').end('down').row
print(row_num)
```

    449
    


```python
data_range = 'A2:GE' + str(row_num)
print(data_range)
```

    A2:GE449
    

## converting excel file into python pandas dataframe
- use options(pd.DataFrame) converter


```python
raw_data = sheet[data_range].options(pd.DataFrame, index=False, header = True).value
raw_data.info()
```

    <class 'pandas.core.frame.DataFrame'>
    RangeIndex: 447 entries, 0 to 446
    Columns: 187 entries, 구분 to 기타지방
    dtypes: object(187)
    memory usage: 653.2+ KB
    


```python
raw_data.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>구분</th>
      <th>전국</th>
      <th>서울</th>
      <th>강북\n14개구</th>
      <th>None</th>
      <th>None</th>
      <th>None</th>
      <th>None</th>
      <th>None</th>
      <th>None</th>
      <th>...</th>
      <th>None</th>
      <th>None</th>
      <th>양산</th>
      <th>거제</th>
      <th>진주</th>
      <th>김해</th>
      <th>통영</th>
      <th>제주도</th>
      <th>제주/\n서귀포</th>
      <th>기타지방</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>강북구</td>
      <td>광진구</td>
      <td>노원구</td>
      <td>도봉구</td>
      <td>동대문구</td>
      <td>마포구</td>
      <td>...</td>
      <td>의창구</td>
      <td>진해구</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>1</th>
      <td>Classification</td>
      <td>Total</td>
      <td>Seoul</td>
      <td>Northern seoul</td>
      <td>Gangbuk-gu</td>
      <td>Gwangjin-gu</td>
      <td>Nowon-gu</td>
      <td>Dobong-gu</td>
      <td>Dongdaemun-gu</td>
      <td>Mapo-gu</td>
      <td>...</td>
      <td>Uichang</td>
      <td>Jinhae</td>
      <td>Yangsan</td>
      <td>Geoje</td>
      <td>Jinju</td>
      <td>Gimhae</td>
      <td>Tongyoung</td>
      <td>Jeju-do</td>
      <td>Jeju/\nSeogwipo</td>
      <td>Non-Metropolitan Area</td>
    </tr>
    <tr>
      <th>2</th>
      <td>86.1</td>
      <td>27.68215</td>
      <td>23.472864</td>
      <td>32.594416</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>3</th>
      <td>2.0</td>
      <td>27.68215</td>
      <td>23.472864</td>
      <td>32.554907</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>4</th>
      <td>3.0</td>
      <td>27.723591</td>
      <td>23.440488</td>
      <td>32.554907</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
  </tbody>
</table>
<p>5 rows × 187 columns</p>
</div>



- The first row and the column name should be merged in the ETL process
- the second row, which is the English name of the corresponding region should be dropped for better readability
- irrelevant rows should be dropped (the last 4 rows)


```python
raw_data.tail(4)
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>구분</th>
      <th>전국</th>
      <th>서울</th>
      <th>강북\n14개구</th>
      <th>None</th>
      <th>None</th>
      <th>None</th>
      <th>None</th>
      <th>None</th>
      <th>None</th>
      <th>...</th>
      <th>None</th>
      <th>None</th>
      <th>양산</th>
      <th>거제</th>
      <th>진주</th>
      <th>김해</th>
      <th>통영</th>
      <th>제주도</th>
      <th>제주/\n서귀포</th>
      <th>기타지방</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>443</th>
      <td>『데이터허브』에서 KB부동산 통계를 편리하게 이용하세요</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>444</th>
      <td>KB부동산 &gt; 메뉴 &gt; 데이터허브</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>445</th>
      <td>데이터허브</td>
      <td>None</td>
      <td>https://data.kbland.kr/</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>446</th>
      <td>주택가격동향</td>
      <td>None</td>
      <td>https://data.kbland.kr/kbstats/wmh?tIdx=HT01&amp;t...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
  </tbody>
</table>
<p>4 rows × 187 columns</p>
</div>




```python
big_col = list(raw_data.columns) # Header column
small_col = list(raw_data.iloc[0]) # the first row
```


```python
bignames = pd.Series(big_col).unique().tolist()
```


```python
bignames.remove(None)
```


```python
for num, gu_data in enumerate(small_col):
    if gu_data == None:
        small_col[num] = big_col[num]
    check = num
    while True:
        if big_col[check] in bignames:
            big_col[num] = big_col[check]
            break
        else:
            check = check - 1
```


```python
raw_data.columns = [big_col, small_col]
```


```python
raw_data.columns
```




    MultiIndex([(      '구분',       '구분'),
                (      '전국',       '전국'),
                (      '서울',       '서울'),
                ('강북\n14개구', '강북\n14개구'),
                ('강북\n14개구',      '강북구'),
                ('강북\n14개구',      '광진구'),
                ('강북\n14개구',      '노원구'),
                ('강북\n14개구',      '도봉구'),
                ('강북\n14개구',     '동대문구'),
                ('강북\n14개구',      '마포구'),
                ...
                (      '창원',      '의창구'),
                (      '창원',      '진해구'),
                (      '양산',       '양산'),
                (      '거제',       '거제'),
                (      '진주',       '진주'),
                (      '김해',       '김해'),
                (      '통영',       '통영'),
                (     '제주도',      '제주도'),
                ('제주/\n서귀포', '제주/\n서귀포'),
                (    '기타지방',     '기타지방')],
               length=187)




```python
raw_data.drop([0,1], inplace=True)
raw_data.drop(raw_data.tail(4).index, inplace=True)
```


```python
raw_data.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead tr th {
        text-align: left;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr>
      <th></th>
      <th>구분</th>
      <th>전국</th>
      <th>서울</th>
      <th colspan="7" halign="left">강북\n14개구</th>
      <th>...</th>
      <th colspan="2" halign="left">창원</th>
      <th>양산</th>
      <th>거제</th>
      <th>진주</th>
      <th>김해</th>
      <th>통영</th>
      <th>제주도</th>
      <th>제주/\n서귀포</th>
      <th>기타지방</th>
    </tr>
    <tr>
      <th></th>
      <th>구분</th>
      <th>전국</th>
      <th>서울</th>
      <th>강북\n14개구</th>
      <th>강북구</th>
      <th>광진구</th>
      <th>노원구</th>
      <th>도봉구</th>
      <th>동대문구</th>
      <th>마포구</th>
      <th>...</th>
      <th>의창구</th>
      <th>진해구</th>
      <th>양산</th>
      <th>거제</th>
      <th>진주</th>
      <th>김해</th>
      <th>통영</th>
      <th>제주도</th>
      <th>제주/\n서귀포</th>
      <th>기타지방</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>2</th>
      <td>86.1</td>
      <td>27.68215</td>
      <td>23.472864</td>
      <td>32.594416</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>3</th>
      <td>2.0</td>
      <td>27.68215</td>
      <td>23.472864</td>
      <td>32.554907</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>4</th>
      <td>3.0</td>
      <td>27.723591</td>
      <td>23.440488</td>
      <td>32.554907</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>5</th>
      <td>4.0</td>
      <td>27.516389</td>
      <td>23.310982</td>
      <td>32.436382</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>6</th>
      <td>5.0</td>
      <td>27.392068</td>
      <td>23.116724</td>
      <td>32.080807</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
  </tbody>
</table>
<p>5 rows × 187 columns</p>
</div>



## Cleansing the dataframe index into timeframe


```python
index_list = list(raw_data['구분']['구분'])

new_index = list()

for num, raw_index in enumerate(index_list):
    temp = str(raw_index).split('.')
    if int(temp[0]) > 12:
        if len(temp[0]) == 2:
            new_index.append('19' + temp[0] + '.' + temp[1])
        else:
            new_index.append(temp[0] + '.' + temp[1])
    else:
        new_index.append(new_index[num-1].split('.')[0] + '.' + temp[0])
```


```python
raw_data.set_index(pd.to_datetime(new_index), inplace=True)
# raw_data.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead tr th {
        text-align: left;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr>
      <th></th>
      <th>구분</th>
      <th>전국</th>
      <th>서울</th>
      <th colspan="7" halign="left">강북\n14개구</th>
      <th>...</th>
      <th colspan="2" halign="left">창원</th>
      <th>양산</th>
      <th>거제</th>
      <th>진주</th>
      <th>김해</th>
      <th>통영</th>
      <th>제주도</th>
      <th>제주/\n서귀포</th>
      <th>기타지방</th>
    </tr>
    <tr>
      <th></th>
      <th>구분</th>
      <th>전국</th>
      <th>서울</th>
      <th>강북\n14개구</th>
      <th>강북구</th>
      <th>광진구</th>
      <th>노원구</th>
      <th>도봉구</th>
      <th>동대문구</th>
      <th>마포구</th>
      <th>...</th>
      <th>의창구</th>
      <th>진해구</th>
      <th>양산</th>
      <th>거제</th>
      <th>진주</th>
      <th>김해</th>
      <th>통영</th>
      <th>제주도</th>
      <th>제주/\n서귀포</th>
      <th>기타지방</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>1986-01-01</th>
      <td>86.1</td>
      <td>27.68215</td>
      <td>23.472864</td>
      <td>32.594416</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>1986-02-01</th>
      <td>2.0</td>
      <td>27.68215</td>
      <td>23.472864</td>
      <td>32.554907</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>1986-03-01</th>
      <td>3.0</td>
      <td>27.723591</td>
      <td>23.440488</td>
      <td>32.554907</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>1986-04-01</th>
      <td>4.0</td>
      <td>27.516389</td>
      <td>23.310982</td>
      <td>32.436382</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>1986-05-01</th>
      <td>5.0</td>
      <td>27.392068</td>
      <td>23.116724</td>
      <td>32.080807</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
  </tbody>
</table>
<p>5 rows × 187 columns</p>
</div>




```python
clean_data = raw_data.drop(('구분','구분'), axis = 1)
```


```python
clean_data
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead tr th {
        text-align: left;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr>
      <th></th>
      <th>전국</th>
      <th>서울</th>
      <th colspan="8" halign="left">강북\n14개구</th>
      <th>...</th>
      <th colspan="2" halign="left">창원</th>
      <th>양산</th>
      <th>거제</th>
      <th>진주</th>
      <th>김해</th>
      <th>통영</th>
      <th>제주도</th>
      <th>제주/\n서귀포</th>
      <th>기타지방</th>
    </tr>
    <tr>
      <th></th>
      <th>전국</th>
      <th>서울</th>
      <th>강북\n14개구</th>
      <th>강북구</th>
      <th>광진구</th>
      <th>노원구</th>
      <th>도봉구</th>
      <th>동대문구</th>
      <th>마포구</th>
      <th>서대문구</th>
      <th>...</th>
      <th>의창구</th>
      <th>진해구</th>
      <th>양산</th>
      <th>거제</th>
      <th>진주</th>
      <th>김해</th>
      <th>통영</th>
      <th>제주도</th>
      <th>제주/\n서귀포</th>
      <th>기타지방</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>1986-01-01</th>
      <td>27.68215</td>
      <td>23.472864</td>
      <td>32.594416</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>1986-02-01</th>
      <td>27.68215</td>
      <td>23.472864</td>
      <td>32.554907</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>1986-03-01</th>
      <td>27.723591</td>
      <td>23.440488</td>
      <td>32.554907</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>1986-04-01</th>
      <td>27.516389</td>
      <td>23.310982</td>
      <td>32.436382</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>1986-05-01</th>
      <td>27.392068</td>
      <td>23.116724</td>
      <td>32.080807</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>...</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
      <td>None</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>2022-05-01</th>
      <td>100.768293</td>
      <td>100.557572</td>
      <td>100.514011</td>
      <td>100.107569</td>
      <td>101.29182</td>
      <td>100.025793</td>
      <td>100.833792</td>
      <td>100.220039</td>
      <td>100.240308</td>
      <td>100.042269</td>
      <td>...</td>
      <td>101.60097</td>
      <td>102.027125</td>
      <td>None</td>
      <td>None</td>
      <td>102.017169</td>
      <td>100.116237</td>
      <td>None</td>
      <td>None</td>
      <td>100.949068</td>
      <td>101.267352</td>
    </tr>
    <tr>
      <th>2022-06-01</th>
      <td>100.868527</td>
      <td>100.723141</td>
      <td>100.673584</td>
      <td>100.170506</td>
      <td>101.431506</td>
      <td>99.931069</td>
      <td>100.927532</td>
      <td>100.279169</td>
      <td>100.498788</td>
      <td>100.622097</td>
      <td>...</td>
      <td>102.668235</td>
      <td>102.383504</td>
      <td>None</td>
      <td>None</td>
      <td>102.557368</td>
      <td>100.108882</td>
      <td>None</td>
      <td>None</td>
      <td>101.229909</td>
      <td>101.491416</td>
    </tr>
    <tr>
      <th>2022-07-01</th>
      <td>100.868825</td>
      <td>100.79044</td>
      <td>100.717396</td>
      <td>100.324139</td>
      <td>101.669564</td>
      <td>99.938522</td>
      <td>100.861091</td>
      <td>100.143208</td>
      <td>100.727687</td>
      <td>100.656617</td>
      <td>...</td>
      <td>102.749177</td>
      <td>102.574134</td>
      <td>None</td>
      <td>None</td>
      <td>102.740061</td>
      <td>100.034374</td>
      <td>None</td>
      <td>None</td>
      <td>101.257225</td>
      <td>101.619213</td>
    </tr>
    <tr>
      <th>2022-08-01</th>
      <td>100.727588</td>
      <td>100.719434</td>
      <td>100.615015</td>
      <td>100.398038</td>
      <td>101.594999</td>
      <td>99.629556</td>
      <td>100.465284</td>
      <td>100.111988</td>
      <td>100.677197</td>
      <td>100.590581</td>
      <td>...</td>
      <td>102.802054</td>
      <td>102.661817</td>
      <td>None</td>
      <td>None</td>
      <td>102.77765</td>
      <td>99.842323</td>
      <td>None</td>
      <td>None</td>
      <td>101.270875</td>
      <td>101.635685</td>
    </tr>
    <tr>
      <th>2022-09-01</th>
      <td>100.56764</td>
      <td>100.6413</td>
      <td>100.50832</td>
      <td>100.419848</td>
      <td>101.684775</td>
      <td>99.09902</td>
      <td>100.270399</td>
      <td>99.968764</td>
      <td>100.688271</td>
      <td>100.680763</td>
      <td>...</td>
      <td>102.743905</td>
      <td>102.726283</td>
      <td>None</td>
      <td>None</td>
      <td>102.611773</td>
      <td>99.621415</td>
      <td>None</td>
      <td>None</td>
      <td>101.351286</td>
      <td>101.591573</td>
    </tr>
  </tbody>
</table>
<p>441 rows × 186 columns</p>
</div>




```python
clean_data.to_csv('output_data.csv', header=True, encoding = 'utf-8-sig')
```


```python

```
