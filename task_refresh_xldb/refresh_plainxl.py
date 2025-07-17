import pandas as pd

# Зчитування обох таблиць
df_main = pd.read_excel('оновити.xlsx')
df_source = pd.read_excel('essentials.xlsx')

# Вказуємо, який стовпець є унікальним для зв'язку
key_col = 'ПІБ'
# Список полів, які треба оновити
columns_to_update = ['конмоб','телефон', 'ддт', 'освіта']

# Перебір рядків у головній таблиці
for index, row in df_main.iterrows():
    key = row[key_col]
    match = df_source[df_source[key_col] == key]
    if not match.empty:
        for col in columns_to_update:
            if pd.isna(row[col]) or str(row[col]).strip() == '':
                new_val = match.iloc[0][col]
                if not pd.isna(new_val):
                    df_main.at[index, col] = new_val

# Збереження
df_main.to_excel('оновлена_таблиця.xlsx', index=False)
