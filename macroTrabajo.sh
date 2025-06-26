if which nix-env &> /dev/null; then
    cd $HOME/Downloads/DriveSyncFiles/NixOs/flakes/MACRO_trabajo;
    # env variables
    source ./env/env.sh
    if python3.10 -c "import pyautogui" &> /dev/null; then
        python3.10 MACRO_entrada-de-materiales.py;
        echo 'si';
    else
        nix develop;
        echo 'no';
    fi
else

if ! command -v nix-env &> /dev/null
then
  echo "Nix no está instalado. Instalando Nix"
  curl --proto '=https' --tlsv1.2 -sSf -L https://install.determinate.systems/nix | sh -s -- install --no-confirm


fi
