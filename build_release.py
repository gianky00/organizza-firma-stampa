import os
import shutil
import subprocess
import sys


def build():
    print("--- Inizio Processo di Build Release ---")

    dist_dir = "dist"
    if os.path.exists(dist_dir):
        print(f"Pulizia directory {dist_dir}...")
        shutil.rmtree(dist_dir)
    os.makedirs(dist_dir)

    # 1. Offuscamento con PyArmor
    print("Offuscamento codice in corso...")
    try:
        # Comando per PyArmor 8.x/9.x
        cmd = [sys.executable, "-m", "pyarmor.cli", "gen", "-O", dist_dir, "-r", "main.py", "src"]
        subprocess.run(cmd, check=True)
        print("Offuscamento completato con successo.")
    except subprocess.CalledProcessError as e:
        print(f"ERRORE durante l'offuscamento: {e}")
        return

    # 2. Copia Asset e file necessari
    print("Copia asset e file di configurazione...")

    assets_src = os.path.join("src", "assets")
    assets_dist = os.path.join(dist_dir, "src", "assets")
    if os.path.exists(assets_src):
        if not os.path.exists(os.path.dirname(assets_dist)):
            os.makedirs(os.path.dirname(assets_dist))
        if os.path.exists(assets_dist):
            shutil.rmtree(assets_dist)
        shutil.copytree(assets_src, assets_dist)
        print("Cartella Assets copiata.")

    # Crea cartelle di lavoro vuote necessarie
    required_dirs = ["FILE EXCEL DA FIRMARE", "PDF", "SCHEDE DA ORGANIZZARE", "SCHEDE ORGANIZZATE", "SCHEDE SENZA DATA"]
    for d in required_dirs:
        path = os.path.join(dist_dir, d)
        os.makedirs(path, exist_ok=True)
        with open(os.path.join(path, ".gitkeep"), "w") as f:
            pass

    print(f"Build completata nella cartella: {dist_dir}")
    print("\nProssimo passo: Compilazione con Inno Setup.")


if __name__ == "__main__":
    build()
