"""
TP : Technologies Multimédias (TMA)
Dr. EL BENANY Med Mahmoud - 18 février 2026
Solutions complètes : TP1, TP2, TP3
"""

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from scipy import signal
from scipy.io import wavfile
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# TP 1
# ============================================================

print("=" * 60)
print("TP 1 - Signaux élémentaires")
print("=" * 60)

# ----------------------------------------------------------
# 1.1 Génération d'un signal sinusoïdal
# ----------------------------------------------------------
f0 = 10          # Fréquence Hz
A  = 1           # Amplitude
phi = 0          # Phase
fs = 100         # Fréquence d'échantillonnage Hz
T  = 1           # Durée en secondes

t = np.arange(0, T, 1/fs)      # Vecteur temps
x = A * np.sin(2 * np.pi * f0 * t + phi)  # Signal sinusoïdal

fig, axes = plt.subplots(3, 1, figsize=(10, 10))
fig.suptitle('TP1 – Signaux élémentaires', fontsize=14, fontweight='bold')

axes[0].plot(t, x, color='steelblue', linewidth=1.5)
axes[0].set_title('1.1 – Signal sinusoïdal x(t) = sin(2π·10·t)')
axes[0].set_xlabel('Temps (s)')
axes[0].set_ylabel('Amplitude')
axes[0].grid(True, alpha=0.4)

# ----------------------------------------------------------
# 1.2 Ajout de bruit et analyse
# ----------------------------------------------------------
np.random.seed(42)
bruit = np.random.randn(len(t))   # Bruit Blanc Gaussien (BBG)
y = x + bruit                     # Signal bruité

axes[1].plot(t, x, color='steelblue', linewidth=1.5, label='Signal pur x(t)')
axes[1].plot(t, y, color='tomato', linewidth=0.8, alpha=0.8, label='Signal bruité y(t)')
axes[1].set_title('1.2 – Signal pur vs Signal bruité (y = x + BBG)')
axes[1].set_xlabel('Temps (s)')
axes[1].set_ylabel('Amplitude')
axes[1].legend()
axes[1].grid(True, alpha=0.4)

# ----------------------------------------------------------
# 1.3 Signaux élémentaires et Convolution
# ----------------------------------------------------------
N = 100
rect = np.zeros(N)
rect[20:41] = 1                   # Signal porte : 1 entre n=20 et n=40

conv_rect = np.convolve(rect, rect)   # Convolution du signal porte avec lui-même
n_conv = np.arange(len(conv_rect))

axes[2].plot(n_conv, conv_rect, color='mediumseagreen', linewidth=1.5)
axes[2].set_title('1.3 – Convolution du signal porte avec lui-même → Signal triangulaire')
axes[2].set_xlabel('Indice n')
axes[2].set_ylabel('Amplitude')
axes[2].grid(True, alpha=0.4)

# Annotation sur la forme
axes[2].annotate('Forme triangulaire\n(convolution de deux portes)', 
                 xy=(40, 20), xytext=(60, 15),
                 arrowprops=dict(arrowstyle='->', color='black'),
                 fontsize=9)

plt.tight_layout()
plt.savefig('/mnt/user-data/outputs/TP1_signaux.png', dpi=150, bbox_inches='tight')
plt.close()
print("→ TP1 sauvegardé : TP1_signaux.png")
print("  La convolution d'un signal porte avec lui-même donne un signal TRIANGULAIRE.")

# ============================================================
# TP 2
# ============================================================

print("\n" + "=" * 60)
print("TP 2 - Analyse spectrale et filtrage")
print("=" * 60)

fs2 = 44100          # Fréquence d'échantillonnage (audio)
duration = 1.0       # 1 seconde
t2 = np.linspace(0, duration, int(fs2 * duration), endpoint=False)

# ----------------------------------------------------------
# 2.1 Signal sonore : f1=440 Hz + f2=880 Hz
# ----------------------------------------------------------
f1, f2 = 440, 880
s_pur = np.sin(2 * np.pi * f1 * t2) + np.sin(2 * np.pi * f2 * t2)

N2 = len(t2)
S_pur = np.fft.fft(s_pur)
freqs = np.fft.fftfreq(N2, 1/fs2)

# Axe positif seulement
half = N2 // 2
freqs_pos = freqs[:half]
S_pos = np.abs(S_pur[:half])

# ----------------------------------------------------------
# 2.2 Ajout bruit fbruit = 5000 Hz
# ----------------------------------------------------------
fbruit = 5000
amplitude_bruit = 0.3
s_bruite = s_pur + amplitude_bruit * np.sin(2 * np.pi * fbruit * t2)

S_bruite = np.fft.fft(s_bruite)
S_bruite_mod = np.abs(S_bruite[:half])

# ----------------------------------------------------------
# 2.3 Filtrage coupe-bande autour de 5000 Hz
# ----------------------------------------------------------
S_filtre = S_bruite.copy()
margin = 200   # ± 200 Hz autour de fbruit

for i, f in enumerate(freqs):
    if fbruit - margin <= abs(f) <= fbruit + margin:
        S_filtre[i] = 0

# 2.4 IFFT → signal filtré
s_filtre = np.real(np.fft.ifft(S_filtre))

# ----------------------------------------------------------
# Tracés TP2
# ----------------------------------------------------------
fig2, axes2 = plt.subplots(2, 2, figsize=(14, 8))
fig2.suptitle('TP2 – Analyse spectrale et filtrage', fontsize=14, fontweight='bold')

# Spectre signal pur
axes2[0, 0].plot(freqs_pos, S_pos, color='steelblue')
axes2[0, 0].set_title('2.1 – Spectre du signal pur (440 + 880 Hz)')
axes2[0, 0].set_xlabel('Fréquence (Hz)')
axes2[0, 0].set_ylabel('|FFT|')
axes2[0, 0].set_xlim(0, 2000)
axes2[0, 0].grid(True, alpha=0.4)
for f_mark in [440, 880]:
    axes2[0, 0].axvline(f_mark, color='red', linestyle='--', alpha=0.7, label=f'{f_mark} Hz')
axes2[0, 0].legend()

# Spectre signal bruité
axes2[0, 1].plot(freqs_pos, S_bruite_mod, color='tomato')
axes2[0, 1].set_title('2.2 – Spectre du signal bruité (pic à 5000 Hz)')
axes2[0, 1].set_xlabel('Fréquence (Hz)')
axes2[0, 1].set_ylabel('|FFT|')
axes2[0, 1].set_xlim(0, fs2//2)
axes2[0, 1].axvline(fbruit, color='purple', linestyle='--', label=f'Bruit {fbruit} Hz')
axes2[0, 1].legend()
axes2[0, 1].grid(True, alpha=0.4)

# Spectre filtré
S_filtre_mod = np.abs(S_filtre[:half])
axes2[1, 0].plot(freqs_pos, S_filtre_mod, color='mediumseagreen')
axes2[1, 0].set_title('2.3 – Spectre après filtrage coupe-bande')
axes2[1, 0].set_xlabel('Fréquence (Hz)')
axes2[1, 0].set_ylabel('|FFT|')
axes2[1, 0].set_xlim(0, fs2//2)
axes2[1, 0].grid(True, alpha=0.4)

# Comparaison temporelle (zoom sur 0.01s)
n_plot = int(0.01 * fs2)
axes2[1, 1].plot(t2[:n_plot]*1000, s_bruite[:n_plot], color='tomato', alpha=0.7, label='Signal bruité')
axes2[1, 1].plot(t2[:n_plot]*1000, s_filtre[:n_plot], color='mediumseagreen', label='Signal filtré')
axes2[1, 1].set_title('2.5 – Comparaison temporelle (10 ms)')
axes2[1, 1].set_xlabel('Temps (ms)')
axes2[1, 1].set_ylabel('Amplitude')
axes2[1, 1].legend()
axes2[1, 1].grid(True, alpha=0.4)

plt.tight_layout()
plt.savefig('/mnt/user-data/outputs/TP2_spectral.png', dpi=150, bbox_inches='tight')
plt.close()
print("→ TP2 sauvegardé : TP2_spectral.png")

# ============================================================
# TP 3
# ============================================================

print("\n" + "=" * 60)
print("TP 3 - Aliasing et Quantification d'image")
print("=" * 60)

# ----------------------------------------------------------
# Partie 1 : Aliasing (simulation avec signal synthétique)
# ----------------------------------------------------------
# On simule un signal audio riche en hautes fréquences
fs_orig = 44100
duration3 = 0.5
t3 = np.linspace(0, duration3, int(fs_orig * duration3), endpoint=False)

# Signal avec plusieurs harmoniques (dont des hautes fréquences)
freqs_test = [440, 1000, 2000, 5000, 8000, 10000]
s_audio = sum(np.sin(2 * np.pi * f * t3) / (i+1) for i, f in enumerate(freqs_test))
s_audio /= np.max(np.abs(s_audio))

# Sous-échantillonnage : garder 1 échantillon sur 10
factor = 10
s_sous = s_audio[::factor]
fs_sous = fs_orig // factor    # 4410 Hz

# Spectrogramme original
fig3, axes3 = plt.subplots(2, 2, figsize=(14, 8))
fig3.suptitle('TP3 – Partie 1 : Aliasing par sous-échantillonnage', fontsize=14, fontweight='bold')

axes3[0, 0].specgram(s_audio, Fs=fs_orig, cmap='inferno')
axes3[0, 0].set_title(f'Spectrogramme original (fs={fs_orig} Hz)')
axes3[0, 0].set_xlabel('Temps (s)')
axes3[0, 0].set_ylabel('Fréquence (Hz)')

axes3[0, 1].specgram(s_sous, Fs=fs_sous, cmap='inferno')
axes3[0, 1].set_title(f'Spectrogramme sous-échantillonné (fs={fs_sous} Hz)')
axes3[0, 1].set_xlabel('Temps (s)')
axes3[0, 1].set_ylabel('Fréquence (Hz)')

# Comparaison spectres FFT
N3_orig = len(s_audio)
S3_orig = np.abs(np.fft.rfft(s_audio))
f3_orig = np.fft.rfftfreq(N3_orig, 1/fs_orig)

N3_sous = len(s_sous)
S3_sous = np.abs(np.fft.rfft(s_sous))
f3_sous = np.fft.rfftfreq(N3_sous, 1/fs_sous)

axes3[1, 0].plot(f3_orig, S3_orig, color='steelblue')
axes3[1, 0].set_title('Spectre original')
axes3[1, 0].set_xlabel('Fréquence (Hz)')
axes3[1, 0].set_ylabel('|FFT|')
axes3[1, 0].axvline(fs_sous/2, color='red', linestyle='--', label='Nyquist sous-ech.')
axes3[1, 0].legend()
axes3[1, 0].grid(True, alpha=0.4)

axes3[1, 1].plot(f3_sous, S3_sous, color='tomato')
axes3[1, 1].set_title('Spectre sous-échantillonné (aliasing visible)')
axes3[1, 1].set_xlabel('Fréquence (Hz)')
axes3[1, 1].set_ylabel('|FFT|')
axes3[1, 1].grid(True, alpha=0.4)

plt.tight_layout()
plt.savefig('/mnt/user-data/outputs/TP3_aliasing.png', dpi=150, bbox_inches='tight')
plt.close()
print("→ TP3 Partie 1 sauvegardée : TP3_aliasing.png")

# ----------------------------------------------------------
# Partie 2 : Quantification et pixelisation d'image
# ----------------------------------------------------------

# Générer une image synthétique en dégradé (simule ciel/peau)
H, W = 256, 512
img = np.zeros((H, W), dtype=np.float64)

# Dégradé horizontal + vertical simulant un ciel avec des détails
x_grid = np.linspace(0, 1, W)
y_grid = np.linspace(0, 1, H)
XX, YY = np.meshgrid(x_grid, y_grid)
img = (0.5 * XX + 0.3 * YY + 0.1 * np.sin(10*XX) + 0.1 * np.cos(8*YY))
img = (img - img.min()) / (img.max() - img.min()) * 255
img = img.astype(np.uint8)

def quantify(image, n_bits):
    """Réduit l'image à 2^n_bits niveaux de gris."""
    n_levels = 2 ** n_bits
    factor = 256 // n_levels
    return (image // factor) * factor

def pixelize(image, block_size):
    """Réduit la résolution spatiale par un facteur block_size."""
    h, w = image.shape
    small = image[::block_size, ::block_size]
    # Agrandir à la taille originale (nearest-neighbor)
    from PIL import Image
    img_small = Image.fromarray(small)
    img_big = img_small.resize((w, h), Image.NEAREST)
    return np.array(img_big)

img_4niveaux = quantify(img, 2)    # 4 niveaux (2 bits)
img_2niveaux = quantify(img, 1)    # 2 niveaux (1 bit)
img_pixel = pixelize(img, 8)       # Pixelisation ÷8

fig4, axes4 = plt.subplots(2, 2, figsize=(12, 8))
fig4.suptitle('TP3 – Partie 2 : Quantification et Pixelisation', fontsize=14, fontweight='bold')

axes4[0, 0].imshow(img, cmap='gray', vmin=0, vmax=255)
axes4[0, 0].set_title('Image originale (8 bits, 256 niveaux)')
axes4[0, 0].axis('off')

axes4[0, 1].imshow(img_4niveaux, cmap='gray', vmin=0, vmax=255)
axes4[0, 1].set_title('Quantification 2 bits (4 niveaux)\n→ Effet "contouring/banding"')
axes4[0, 1].axis('off')

axes4[1, 0].imshow(img_2niveaux, cmap='gray', vmin=0, vmax=255)
axes4[1, 0].set_title('Quantification 1 bit (2 niveaux)\n→ Image binaire, banding sévère')
axes4[1, 0].axis('off')

axes4[1, 1].imshow(img_pixel, cmap='gray', vmin=0, vmax=255)
axes4[1, 1].set_title('Pixelisation (÷8, affiché à taille originale)\n→ Effet mosaïque')
axes4[1, 1].axis('off')

plt.tight_layout()
plt.savefig('/mnt/user-data/outputs/TP3_quantification.png', dpi=150, bbox_inches='tight')
plt.close()
print("→ TP3 Partie 2 sauvegardée : TP3_quantification.png")

# ============================================================
# Résumé des réponses aux questions théoriques
# ============================================================
print("\n" + "=" * 60)
print("RÉSUMÉ DES RÉPONSES AUX QUESTIONS")
print("=" * 60)
print("""
TP1 – 1.3 : Forme du signal résultant
  La convolution d'un signal porte avec lui-même donne un signal
  TRIANGULAIRE. C'est une propriété fondamentale : rect ★ rect = triangle.

TP3 – Partie 1 : Observation sur les hautes fréquences
  Après sous-échantillonnage (fs/10 = 4410 Hz), toute fréquence
  supérieure à 2205 Hz (= fs_sous/2, limite de Nyquist) est REPLIÉE
  (aliasing). Les sons aigus et consonnes sifflantes deviennent des
  artefacts à basses fréquences inexistants dans le signal original.

TP3 – Partie 2 : Effet visuel de la quantification
  En réduisant à 4 ou 2 niveaux, les zones de dégradé (ciel, peau)
  présentent des "sauts" brusques entre niveaux → effet appelé
  CONTOURING ou BANDING (contourage / effet de bande).
  La pixelisation ÷8 produit un effet MOSAÏQUE (blocs carrés visibles).
""")