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
plt.savefig('TP/TP1/TP1_signaux.png', dpi=150, bbox_inches='tight')
plt.close()
print("→ TP1 sauvegardé : TP1_signaux.png")
print("  La convolution d'un signal porte avec lui-même donne un signal TRIANGULAIRE.")

