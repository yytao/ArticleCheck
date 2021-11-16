import os

path = os.path.abspath(os.path.dirname(__file__))
print(path)
with open(path+'\\keyword.config', encoding='utf8') as f:
    data = f.read()
    print(data)


    