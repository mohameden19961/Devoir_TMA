# ============================================================
# TP 3 - Aliasing et Quantification d'image
# Version finale pour remise
# ============================================================

import numpy as np
import matplotlib.pyplot as plt
from PIL import Image

print("\n" + "=" * 60)
print("TP 3 - Aliasing et Quantification d'image (Final)")
print("=" * 60)

# ============================================================
# Partie 1 : Aliasing (signal synthétique)
# ============================================================

# 1️⃣ Paramètres signal
fs_orig = 44100       # Fréquence échantillonnage originale
duration3 = 0.5       # Durée 0.5 s
t3 = np.linspace(0, duration3, int(fs_orig*duration3), endpoint=False)

# 2️⃣ Signal avec plusieurs harmoniques (riches en hautes fréquences)
freqs_test = [440, 1000, 2000, 5000, 8000, 10000]
s_audio = sum(np.sin(2*np.pi*f*t3)/(i+1) for i,f in enumerate(freqs_test))
s_audio /= np.max(np.abs(s_audio))   # normalisation

# 3️⃣ Sous-échantillonnage pour créer aliasing
factor = 10
s_sous = s_audio[::factor]
fs_sous = fs_orig // factor  # 4410 Hz

# 4️⃣ FFT pour comparaison spectres
N3_orig = len(s_audio)
S3_orig = np.abs(np.fft.rfft(s_audio))
f3_orig = np.fft.rfftfreq(N3_orig, 1/fs_orig)

N3_sous = len(s_sous)
S3_sous = np.abs(np.fft.rfft(s_sous))
f3_sous = np.fft.rfftfreq(N3_sous, 1/fs_sous)

# 5️⃣ Tracés spectrogrammes et FFT
fig3, axes3 = plt.subplots(2,2, figsize=(14,8))
fig3.suptitle("TP3 – Partie 1 : Aliasing par sous-échantillonnage", fontsize=14, fontweight='bold')

# Spectrogramme original
axes3[0,0].specgram(s_audio, Fs=fs_orig, cmap='inferno')
axes3[0,0].set_title(f"Spectrogramme original (fs={fs_orig} Hz)")
axes3[0,0].set_xlabel("Temps (s)")
axes3[0,0].set_ylabel("Fréquence (Hz)")

# Spectrogramme sous-échantillonné
axes3[0,1].specgram(s_sous, Fs=fs_sous, cmap='inferno')
axes3[0,1].set_title(f"Spectrogramme sous-échantillonné (fs={fs_sous} Hz)")
axes3[0,1].set_xlabel("Temps (s)")
axes3[0,1].set_ylabel("Fréquence (Hz)")

# Spectre FFT original
axes3[1,0].plot(f3_orig, S3_orig, color='steelblue')
axes3[1,0].set_title("Spectre original")
axes3[1,0].set_xlabel("Fréquence (Hz)")
axes3[1,0].set_ylabel("|FFT|")
axes3[1,0].axvline(fs_sous/2, color='red', linestyle='--', label='Nyquist sous-échantillonné')
axes3[1,0].legend()
axes3[1,0].grid(True, alpha=0.4)

# Spectre FFT sous-échantillonné
axes3[1,1].plot(f3_sous, S3_sous, color='tomato')
axes3[1,1].set_title("Spectre sous-échantillonné (aliasing visible)")
axes3[1,1].set_xlabel("Fréquence (Hz)")
axes3[1,1].set_ylabel("|FFT|")
axes3[1,1].grid(True, alpha=0.4)

plt.tight_layout()
plt.savefig("TP3_aliasing_final.png", dpi=150)
plt.close()
print("→ Partie 1 sauvegardée : TP3_aliasing_final.png")

# ============================================================
# Partie 2 : Quantification et Pixelisation d'image
# ============================================================

# 1️⃣ Générer image synthétique
H, W = 256, 512
x_grid = np.linspace(0,1,W)
y_grid = np.linspace(0,1,H)
XX, YY = np.meshgrid(x_grid, y_grid)
img = (0.5*XX + 0.3*YY + 0.1*np.sin(10*XX) + 0.1*np.cos(8*YY))
img = (img - img.min()) / (img.max()-img.min()) * 255
img = img.astype(np.uint8)

# 2️⃣ Fonctions utilitaires
def quantify(image, n_bits):
    """Quantifie l'image à 2^n_bits niveaux."""
    n_levels = 2 ** n_bits
    factor = 256 // n_levels
    return (image // factor) * factor

def pixelize(image, block_size):
    """Pixelisation avec nearest neighbor."""
    small = image[::block_size, ::block_size]
    img_small = Image.fromarray(small)
    img_big = img_small.resize((image.shape[1], image.shape[0]), Image.NEAREST)
    return np.array(img_big)

# 3️⃣ Quantification et pixelisation
img_4niveaux = quantify(img, 2)  # 4 niveaux (2 bits)
img_2niveaux = quantify(img, 1)  # 2 niveaux (1 bit)
img_pixel = pixelize(img, 8)     # Pixelisation ÷8

# 4️⃣ Tracés images
fig4, axes4 = plt.subplots(2,2, figsize=(12,8))
fig4.suptitle("TP3 – Partie 2 : Quantification et Pixelisation", fontsize=14, fontweight='bold')

axes4[0,0].imshow(img, cmap='gray', vmin=0, vmax=255)
axes4[0,0].set_title("Image originale (8 bits, 256 niveaux)")
axes4[0,0].axis("off")

axes4[0,1].imshow(img_4niveaux, cmap='gray', vmin=0, vmax=255)
axes4[0,1].set_title("Quantification 2 bits (4 niveaux)")
axes4[0,1].axis("off")

axes4[1,0].imshow(img_2niveaux, cmap='gray', vmin=0, vmax=255)
axes4[1,0].set_title("Quantification 1 bit (2 niveaux)")
axes4[1,0].axis("off")

axes4[1,1].imshow(img_pixel, cmap='gray', vmin=0, vmax=255)
axes4[1,1].set_title("Pixelisation (÷8, nearest neighbor)")
axes4[1,1].axis("off")

plt.tight_layout()
plt.savefig("TP/TP3/TP3_quantification_final.png", dpi=150)
plt.close()
print("→ Partie 2 sauvegardée : TP3_quantification.png")

# ============================================================
# Résumé : le TP est complet
# ============================================================
print("\nTP3 final prêt pour remise : Aliasing et Quantification terminés ✅")