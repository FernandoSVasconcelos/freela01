def get_valores(classe):
    print("Acessando os valores do excel")
    print(f"Formato do canal: {classe.Var1.get()}")
    print(f"Vazão: {classe.vazão}")
    print(f"Rugosidade: {classe.rugosidade}")
    print(f'Teta: {classe.teta}')
    print(f"d: {classe.d}")
    print(f"Base menor: {classe.base_menor}")
    print(f"z: {classe.z}")