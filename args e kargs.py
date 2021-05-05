def argo (*args):
    for i in args:
        print(i)

argo(1, 2, 3, 52)


def kargo (**kwargs):
    x1 = kwargs.get('x')  # melhor solução para o caso de uso do Kwrgs pois a não existência do dados e retornado NONE
    x2 = kwargs.get('b')
    x3 = kwargs['x']  # possivel a utilização desta forma, entretanto, a não existência do dados retorna um erro
    x4 = kwargs['b']

    print(x1,x2,x3,x4)

kargo(x=2,b=69)