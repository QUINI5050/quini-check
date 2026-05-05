def revisar_premios(jugadas, resultados):
    detalle_completo = []
    for jugada in jugadas:
        fila = {"nombre": jugada["nombre"], "modalidades": {}}
        for modalidad, nums in resultados.items():
            aciertos = len(set(jugada["numeros"]) & set(nums))
            fila["modalidades"][modalidad] = {
                "aciertos": aciertos,
                "numeros_jugados": sorted(jugada["numeros"]),
                "numeros_sorteados": sorted(nums)
            }
        detalle_completo.append(fila)
    return detalle_completo