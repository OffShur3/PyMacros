{
  description = "Flake del Trabajo";

  inputs = {
    nixpkgs.url = "github:nixos/nixpkgs/nixos-unstable";
  };
  outputs = {
    self,
    nixpkgs,
  }: let
    system = "x86_64-linux";
    pkgs = nixpkgs.legacyPackages.${system};
  in {
    devShells.${system}.default =
      pkgs.mkShell
      {
        buildInputs = [
          (pkgs.python312.withPackages (ps: with ps; [
            python-dotenv
            pyautogui
            numpy
            pillow
          ]))
        ];
        shellHook = ''
          echo   ""
          printf "\e[3;34m-- Bienvenido a la flake del Trabajo --\e[0m\n"
          echo   ""
          python3 ./MACRO_entrada-de-materiales.py
        '';
      };
  };
}
