import subprocess
from subprocess import check_output
from unittest import result


nombreEntorno = "venv-nomina"
ejecutarVirtualEnv = nombreEntorno + "\\Scripts\\activate.bat &&"
ejecutarPython = "python main.py"

comandos = ejecutarVirtualEnv + ejecutarPython

resultado = subprocess.run(comandos, shell=True)
resultado.check_returncode()

input("Ha finalizado la ejecución \n " +
      " Preciona (ENTER) para continuar... ")