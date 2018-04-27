

def test(n=1):
    for i in range(n):
        print('test', end='')


dic = {
    '1': 'string',
    '2': 123,
    '3': [1, 2, 3],
    '4': test,
}

k = '4'

if isinstance(dic[k], type(test)):
    dic[k](3)
else:
    print(dic[k])



for k, v in dic.items():
    print(k, v)