import matplotlib.pyplot as plt

# Total CTB oil profile
plt.figure(figsize=(12, 5))

plt.plot(
    aggregated["Date"],
    aggregated["Oil_BPD"]
)

plt.title("Silvertip CTB - Total Oil Profile")
plt.xlabel("Date")
plt.ylabel("Oil Rate (BPD)")
plt.grid(True, alpha=0.3)

plt.tight_layout()
plt.show()

plt.figure(figsize=(12, 5))

plt.plot(
    aggregated["Date"],
    aggregated["Gas_MCFD"]
)

plt.title("Silvertip CTB - Total Gas Profile")
plt.xlabel("Date")
plt.ylabel("Gas Rate (MCF/D)")
plt.grid(True, alpha=0.3)

plt.tight_layout()
plt.show()
