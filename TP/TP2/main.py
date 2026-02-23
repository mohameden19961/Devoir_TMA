# ============================================================
# TP 2 - Analyse spectrale et filtrage
# Version finale pour remise
# ============================================================

import numpy as np
import matplotlib.pyplot as plt
from scipy.io import wavfile

print("\n" + "=" * 60)
print("TP 2 - Analyse spectrale et filtrage (Final)")
print("=" * 60)

# ----------------------------------------------------------
# Paramètres généraux
# ----------------------------------------------------------
fs = 44100           # Fréquence d'échantillonnage audio (Hz)
duration = 2.0       # Durée du signal en secondes
t = np.linspace(0, duration, int(fs*duration), endpoint=False)

# ----------------------------------------------------------
# Fonction utilitaire : sauvegarde WAV robuste
# ----------------------------------------------------------
def save_audio(filename, sig, fs):
    """Normalise et sauvegarde un signal en WAV 16-bit."""
    if np.max(np.abs(sig)) > 0:
        sig_norm = sig / np.max(np.abs(sig))
    else:
        sig_norm = sig
    wavfile.write(filename, fs, (sig_norm * 32767).astype(np.int16))
    print(f"Fichier sauvegardé : {filename}")

# ============================================================
# 2.1 – Analyse spectrale d’un signal pur
# ============================================================

# 1️⃣ Génération du signal pur : f1=440 Hz, f2=880 Hz
f1, f2 = 440, 880
signal_pur = np.sin(2*np.pi*f1*t) + np.sin(2*np.pi*f2*t)

# 2️⃣ FFT du signal pur
N = len(signal_pur)
S_pur = np.fft.fft(signal_pur)
freqs = np.fft.fftfreq(N, 1/fs)

# Axe positif seulement
half = N//2
freqs_pos = freqs[:half]
S_pur_mod = np.abs(S_pur[:half])/N

# ============================================================
# 2.2 – Identification et filtrage d’un bruit
# ============================================================

# 1️⃣ Ajout du bruit : fbruit = 5000 Hz
fbruit = 5000
amp_bruit = 0.5
signal_bruite = signal_pur + amp_bruit*np.sin(2*np.pi*fbruit*t)

# FFT du signal bruité
S_bruite = np.fft.fft(signal_bruite)
S_bruite_mod = np.abs(S_bruite[:half])/N

# 2️⃣ Filtrage coupe-bande autour du bruit
S_filtre = S_bruite.copy()
margin = 100  # ±100 Hz autour du bruit
mask = (np.abs(freqs) >= fbruit-margin) & (np.abs(freqs) <= fbruit+margin)
S_filtre[mask] = 0

# Reconstruction temporelle du signal filtré
signal_filtre = np.real(np.fft.ifft(S_filtre))
S_filtre_mod = np.abs(S_filtre[:half])/N

# ============================================================
# 3️⃣ Sauvegarde des fichiers audio
# ============================================================
save_audio("TP/TP2/signal_pur.wav", signal_pur, fs)
save_audio("TP/TP2/signal_bruite.wav", signal_bruite, fs)
save_audio("TP/TP2/signal_filtre.wav", signal_filtre, fs)

# ============================================================
# 4️⃣ Tracés graphiques
# ============================================================
fig, axes = plt.subplots(2, 2, figsize=(14,10))
fig.suptitle("TP2 – Analyse Spectrale et Filtrage", fontsize=16, fontweight='bold')

# ----- Signal pur
axes[0,0].plot(freqs_pos, S_pur_mod, color='blue')
axes[0,0].set_title("2.1 – Spectre du signal pur (440 + 880 Hz)")
axes[0,0].set_xlim(0,1500)
axes[0,0].set_xlabel("Fréquence (Hz)")
axes[0,0].set_ylabel("|FFT|")
axes[0,0].grid(True)
axes[0,0].axvline(f1, color='red', linestyle='--', alpha=0.7, label='f1=440 Hz')
axes[0,0].axvline(f2, color='orange', linestyle='--', alpha=0.7, label='f2=880 Hz')
axes[0,0].legend()

# ----- Signal bruité
axes[0,1].plot(freqs_pos, S_bruite_mod, color='red')
axes[0,1].set_title("2.2 – Spectre du signal bruité (pic à 5000 Hz)")
axes[0,1].set_xlim(0,6000)
axes[0,1].set_xlabel("Fréquence (Hz)")
axes[0,1].set_ylabel("|FFT|")
axes[0,1].grid(True)
axes[0,1].axvline(fbruit, color='purple', linestyle='--', alpha=0.7, label='Bruit 5000 Hz')
axes[0,1].legend()

# ----- Signal filtré
axes[1,0].plot(freqs_pos, S_filtre_mod, color='green')
axes[1,0].set_title("Spectre filtré (coupe-bande 5000 Hz)")
axes[1,0].set_xlim(0,6000)
axes[1,0].set_xlabel("Fréquence (Hz)")
axes[1,0].set_ylabel("|FFT|")
axes[1,0].grid(True)

# ----- Comparaison temporelle (zoom 20 ms)
n_plot = int(0.02*fs)
axes[1,1].plot(t[:n_plot]*1000, signal_bruite[:n_plot], label='Signal bruité', alpha=0.7, color='red')
axes[1,1].plot(t[:n_plot]*1000, signal_filtre[:n_plot], label='Signal filtré', color='blue')
axes[1,1].set_title("Comparaison temporelle (Zoom 20 ms)")
axes[1,1].set_xlabel("Temps (ms)")
axes[1,1].set_ylabel("Amplitude")
axes[1,1].legend()
axes[1,1].grid(True)

plt.tight_layout(rect=[0,0.03,1,0.95])
plt.savefig("TP/TP2/TP2_spectral.png", dpi=150)
plt.show()

print("Image sauvegardée : TP2_spectral_final.png")