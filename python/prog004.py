def break_teste():
    while True:
        s = input('Escreva algo: ')
        if s == 'quit':
            break
        print('Contagem: ', len(s))
    print('Pronto!')

break_teste()
