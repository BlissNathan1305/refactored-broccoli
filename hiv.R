# -----------------------------
# Simulate HIV prevalence data
# -----------------------------
age_groups <- c("15-19", "20-24", "25-29", "30-34")
men <- c(2.1, 3.5, 4.2, 5.0)    # % HIV prevalence among young men
women <- c(3.0, 5.0, 6.0, 6.5)  # % HIV prevalence among young women

# Combine into a data frame
hiv_data <- data.frame(
  Age = age_groups,
  Men = men,
  Women = women
)

print(hiv_data)

# -----------------------------
# Save plot as JPEG
# -----------------------------
jpeg(filename = "hiv_prevalence_nigeria.jpeg", width = 1000, height = 600)

# Set up plot area
barplot(
  t(as.matrix(hiv_data[,2:3])),   # Transpose for grouped bars
  beside = TRUE,                  # Separate bars
  col = c("skyblue", "pink"),     # Colors for men and women
  names.arg = hiv_data$Age,
  main = "HIV Prevalence among Young Men and Women in Nigeria",
  xlab = "Age Group",
  ylab = "HIV Prevalence (%)",
  ylim = c(0, max(hiv_data[,2:3]) + 2)
)
legend("topright", legend = c("Men", "Women"), fill = c("skyblue", "pink"))

# Close JPEG device
dev.off()

cat("HIV prevalence plot saved as hiv_prevalence_nigeria.jpeg\n")
