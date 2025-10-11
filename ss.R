# Generate numbers
x <- 1:10
y <- x^2

# Print the numbers
print(data.frame(x, y))

# Save plot as PNG
png(filename = "simple_plot.png", width = 800, height = 600)

# Plot a simple graph
plot(x, y,
     type = "b",       # points and lines
     col = "blue",     # line color
     pch = 19,         # solid points
     main = "Simple x vs x^2 Plot",
     xlab = "X values",
     ylab = "Y = X^2")

# Close the PNG device
dev.off()

cat("Plot saved as simple_plot.png\n")
