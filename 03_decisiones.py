# ===========================================
# Lección 3: Decisiones (if / else)
# ===========================================
#
# A veces queremos que el programa haga cosas
# diferentes según una condición.
# Es como decir: "SI llueve, llevo paraguas,
#                  SI NO, llevo lentes de sol"

temperatura = 35

if temperatura > 30:
    print("¡Hace mucho calor! Toma agua.")
elif temperatura > 20:
    print("El clima está agradable.")
else:
    print("Hace frío, lleva abrigo.")

# -------------------------------------------
# Explicación:
#   if    = SI (la primera condición)
#   elif  = SI NO, PERO SI (otra condición)
#   else  = SI NADA DE LO ANTERIOR (el caso por defecto)
#
# Los símbolos de comparación:
#   >   mayor que
#   <   menor que
#   ==  igual a (¡ojo! son DOS signos de igual)
#   !=  diferente de
#   >=  mayor o igual
#   <=  menor o igual
# -------------------------------------------

# Otro ejemplo: verificar la edad

edad = 17

if edad >= 18:
    print("Eres mayor de edad")
else:
    print("Eres menor de edad")

# -------------------------------------------
# Ejercicio: Cambia el valor de 'temperatura' y 'edad'
# y observa cómo cambia el resultado.
# Ejecuta con: python3 03_decisiones.py
# -------------------------------------------
