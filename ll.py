from wordcloud import WordCloud
import matplotlib.pyplot as plt

# List of 50 engineering terms
engineering_terms = [
    "thermodynamics", "fluid", "mechanics", "circuit", "voltage", "current", "resistance", "inductor", "capacitor",
    "transistor", "semiconductor", "algorithm", "data", "structure", "compiler", "network", "protocol", "signal",
    "control", "robotics", "automation", "design", "CAD", "stress", "strain", "load", "beam", "torque", "gear",
    "motor", "sensor", "actuator", "feedback", "system", "modeling", "simulation", "analysis", "efficiency",
    "power", "energy", "heat", "transfer", "material", "composite", "welding", "fabrication", "manufacturing",
    "process", "quality", "safety", "engineering"
]

# Combine terms into a single string
text = " ".join(engineering_terms)

# Generate word cloud
wordcloud = WordCloud(width=800, height=400, background_color='white').generate(text)

# Save as JPEG
wordcloud.to_file("engineering_wordcloud.jpeg")

# Display the image
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis("off")
plt.show()
