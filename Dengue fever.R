


###### Automation script 

### Denge was used in this example but the script can be adapted to any other disease.

#### the script is generating a Weekly Trend Lines for Dengue for All Quarters (2024)

### it is also generating a degmographic analytcics and graphs for gender, nationalites, and origin
((( ### it can be adapted to the requried vairables you have in your dataset)))






########## Section 1: Weekly Trend Lines for Dengue for All Quarters ###########

# Load necessary libraries
library(readxl)
library(dplyr)
library(ggplot2)
library(scales)
library(writexl)

# Set file path for the dataset
file_path <- "/Users/turkialmalki/Desktop/GCC weekly data/merged_data_corrected.xlsx"

# Read the dengue sheet
dengue_data <- read_excel(file_path, sheet = "Dengue")

# Filter for 2024 data
dengue_2024 <- dengue_data %>%
  filter(year(`Starting date of Epidemiological week`) == 2024) %>%
  mutate(week = isoweek(as.Date(`Starting date of Epidemiological week`)))

# Manually add population for each country (latest provided values)
population_data <- data.frame(
  Country = c("Bahrain", "Kuwait", "Oman", "Qatar", "Saudi Arabia", "UAE"),
  Population = c(1524693, 4793568, 4933850, 3098866, 32175224, 10032213)
)

# Merge population data
dengue_2024 <- dengue_2024 %>%
  left_join(population_data, by = "Country")

# Calculate incidence rate
dengue_2024 <- dengue_2024 %>%
  group_by(Country, week) %>%
  summarise(
    total_cases = sum(`Weekly cases`, na.rm = TRUE),
    population = first(Population),
    incidence_rate = (total_cases / population) * 100000,
    .groups = "drop"
  )

# Save processed data as an Excel file
output_folder <- "/Users/turkialmalki/Desktop/GCC weekly data/Dengue Graph T"
if (!dir.exists(output_folder)) {
  dir.create(output_folder, recursive = TRUE)
}
write_xlsx(dengue_2024, file.path(output_folder, "Dengue_Calculated_Data.xlsx"))

cat("Processed data has been saved to:", file.path(output_folder, "Dengue_Calculated_Data.xlsx"), "\n")

# Set color codes for countries
country_colors <- c(
  "Bahrain" = "#AB0C0C",
  "Kuwait" = "#333333",
  "Oman" = "#AAD99E",
  "Qatar" = "#701D46",
  "Saudi Arabia" = "#009A4D",
  "UAE" = "#cccccc"
)

# Helper function to generate month labels for x-axis
get_month_labels <- function(year) {
  month_starts <- seq.Date(
    from = as.Date(paste0(year, "-01-01")),
    to = as.Date(paste0(year, "-12-01")),
    by = "month"
  )
  month_labels <- data.frame(
    month_name = substr(month.name[month(month_starts)], 1, 3),
    week = isoweek(month_starts)
  )
  return(month_labels)
}

# Generate month labels for 2024
month_labels <- get_month_labels(2024)

# Plot total cases trend
total_cases_plot <- ggplot(dengue_2024, aes(x = week, y = total_cases, color = Country, group = Country)) +
  geom_line(size = 0.5) +
  geom_point(size = 1) +
  scale_color_manual(values = country_colors) +
  guides(color = guide_legend(
    override.aes = list(shape = 21, size = 4, fill = country_colors[levels(factor(dengue_2024$Country))], linetype = 0)
  )) +
  geom_vline(xintercept = month_labels$week, color = "#E6E6E6", linetype = "dashed", size = 0.3) +
  scale_x_continuous(
    limits = c(1, 53),
    breaks = seq(0, 52, by = 2),
    expand = c(0, 0),
    sec.axis = sec_axis(~ ., breaks = month_labels$week, labels = month_labels$month_name)
  ) +
  labs(
    title = "Weekly Trend of Total Dengue Cases by Country - 2024",
    x = "Epidemiological Week",
    y = "Number of Cases",
    color = "Country"
  ) +
  theme_classic(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, face = "bold", hjust = 0.5),
    axis.title.x = element_text(size = 12, face = "bold"),
    axis.title.x.top = element_text(size = 10),
    axis.title.y = element_text(size = 12, face = "bold"),
    axis.text.x = element_text(size = 10),
    axis.text.x.top = element_text(size = 10, vjust = 0.8, margin = margin(t = -5)),
    axis.text.y = element_text(size = 12),
    axis.ticks.x.top = element_blank()
  )

# Plot incidence rate trend
incidence_rate_plot <- ggplot(dengue_2024, aes(x = week, y = incidence_rate, color = Country, group = Country)) +
  geom_line(size = 0.5) +
  geom_point(size = 1) +
  scale_color_manual(values = country_colors) +
  guides(color = guide_legend(
    override.aes = list(shape = 21, size = 4, fill = country_colors[levels(factor(dengue_2024$Country))], linetype = 0)
  )) +
  geom_vline(xintercept = month_labels$week, color = "#E6E6E6", linetype = "dashed", size = 0.3) +
  scale_x_continuous(
    limits = c(1, 53),
    breaks = seq(0, 52, by = 2),
    expand = c(0, 0),
    sec.axis = sec_axis(~ ., breaks = month_labels$week, labels = month_labels$month_name)
  ) +
  labs(
    title = "Weekly Trend of Dengue Incidence Rate by Country - 2024",
    x = "Epidemiological Week",
    y = "Incidence Rate/100,000 Individuals",
    color = "Country"
  ) +
  theme_classic(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, face = "bold", hjust = 0.5),
    axis.title.x = element_text(size = 12, face = "bold"),
    axis.title.x.top = element_text(size = 10),
    axis.title.y = element_text(size = 12, face = "bold"),
    axis.text.x = element_text(size = 10),
    axis.text.x.top = element_text(size = 10, vjust = 0.8, margin = margin(t = -5)),
    axis.text.y = element_text(size = 12),
    axis.ticks.x.top = element_blank()
  )

### 🔹 NEW SECTION: PRINT PLOTS IN R ###
print(total_cases_plot)  # Display total cases trend in R
print(incidence_rate_plot)  # Display incidence rate trend in R

# Save the charts
ggsave(file.path(output_folder, "weekly_total_cases_trend.png"), total_cases_plot, width = 12, height = 8, dpi = 300)
ggsave(file.path(output_folder, "weekly_incidence_rate_trend.png"), incidence_rate_plot, width = 12, height = 8, dpi = 300)

cat("Trend line charts and Excel file saved successfully in:", output_folder)






####------------------------------------------------------------------------------------------------------####
######Section 2: Gender, Nationality, and Origin Stacked Bar Charts#######

# Load necessary libraries
library(readxl)
library(dplyr)
library(ggplot2)
library(scales)
library(tidyr)

# File path and sheet
file_path <- "/Users/turkialmalki/Desktop/GCC weekly data/merged_data_corrected.xlsx"
sheet_name <- "Dengue"

# Read the dengue sheet
dengue_data <- read_excel(file_path, sheet = sheet_name)

# Calculate the quarter from epidemiological weeks
dengue_data <- dengue_data %>%
  mutate(
    Date = as.Date(`Starting date of Epidemiological week`),
    year = year(Date),
    week = isoweek(Date),
    quarter = case_when(
      week >= 1 & week <= 13 ~ 1,
      week >= 14 & week <= 26 ~ 2,
      week >= 27 & week <= 39 ~ 3,
      week >= 40 & week <= 53 ~ 4
    )
  )

# Generate aggregated datasets for gender, nationality, and origin
gender_data <- dengue_data %>%
  group_by(Country, quarter) %>%
  summarise(
    Male = sum(Male, na.rm = TRUE),
    Female = sum(Female, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  pivot_longer(cols = c(Male, Female), names_to = "Gender", values_to = "Cases") %>%
  mutate(Gender = factor(Gender, levels = c("Female", "Male")))

nationality_data <- dengue_data %>%
  group_by(Country, quarter) %>%
  summarise(
    National = sum(National, na.rm = TRUE),
    `Non-National` = sum(`Non-National`, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  pivot_longer(cols = c(National, `Non-National`), names_to = "Nationality", values_to = "Cases") %>%
  mutate(Nationality = factor(Nationality, levels = c("Non-National", "National")))

origin_data <- dengue_data %>%
  group_by(Country, quarter) %>%
  summarise(
    Imported = sum(Imported, na.rm = TRUE),
    Autochthonous = sum(`Not Imported`, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  pivot_longer(cols = c(Imported, Autochthonous), names_to = "Origin", values_to = "Cases") %>%
  mutate(Origin = factor(Origin, levels = c("Autochthonous", "Imported")))

# Country color mapping
country_colors <- c(
  "Bahrain" = "#AB0C0C",
  "Kuwait" = "#333333",
  "Oman" = "#AAD99E",
  "Qatar" = "#701D46",
  "Saudi Arabia" = "#009A4D",
  "UAE" = "#cccccc"
)

# Function to create faceted 100% stacked bar charts with crosshatch patterns and styled labels
plot_stacked_bars_adjusted <- function(data, fill_var, title) {
  data <- data %>%
    filter(!is.na(Cases) & Cases > 0) %>%  # Exclude missing or zero values
    group_by(Country, quarter) %>%
    mutate(
      percentage = Cases / sum(Cases),  # Calculate percentage
      cumulative_percentage = cumsum(percentage),  # Cumulative percentage for label positioning
      label_position = cumulative_percentage - (percentage / 2)  # Position labels in the middle of the segment
    )
  
  ggplot(data, aes(x = Country, y = percentage, fill = Country, pattern = !!sym(fill_var))) +
    geom_bar_pattern(
      stat = "identity", position = "fill", width = 0.4,  # Narrower bars
      pattern_color = "white", pattern_fill = NA,
      pattern_angle = 45, pattern_density = 0.5, pattern_spacing = 0.02
    ) +
    scale_y_continuous(labels = scales::percent_format(scale = 100)) +
    scale_fill_manual(values = country_colors) +
    scale_pattern_manual(values = c(
      "Male" = "none", "Female" = "crosshatch",
      "National" = "none", "Non-National" = "crosshatch",
      "Imported" = "crosshatch", "Autochthonous" = "none"
    )) +
    facet_wrap(~quarter, ncol = 1, strip.position = "top", labeller = as_labeller(c("1" = "Q1", "2" = "Q2", "3" = "Q3", "4" = "Q4"))) +
    labs(
      title = title,
      x = "Country",
      y = "Percentage"
    ) +
    geom_label(
      aes(label = paste0(!!sym(fill_var), "\n(n=", Cases, ")"), y = label_position),
      color = "black", fill = "white", fontface = "bold", size = 3, label.size = 0.3  # White textbox with black text
    ) +
    theme_classic(base_size = 14) +
    theme(
      plot.title = element_text(size = 18, face = "bold", hjust = 0.5),
      axis.title.x = element_text(size = 12, face = "bold"),
      axis.title.y = element_text(size = 12, face = "bold"),
      axis.text.x = element_text(size = 10, angle = 45, hjust = 1),
      axis.text.y = element_text(size = 10),
      strip.text = element_text(size = 12, face = "bold"),
      legend.position = "none"  # Remove legends
    )
}

# Generate adjusted charts
gender_chart <- plot_stacked_bars_adjusted(
  data = gender_data,
  fill_var = "Gender",
  title = "Gender Distribution of Dengue Cases by Quarter - 2024"
)

nationality_chart <- plot_stacked_bars_adjusted(
  data = nationality_data,
  fill_var = "Nationality",
  title = "Nationality Distribution of Dengue Cases by Quarter - 2024"
)

origin_chart <- plot_stacked_bars_adjusted(
  data = origin_data,
  fill_var = "Origin",
  title = "Origin Distribution of Dengue Cases by Quarter - 2024"
)

# Save the adjusted charts
output_folder <- "/Users/turkialmalki/Desktop/GCC weekly data/Dengue Graph T"
if (!dir.exists(output_folder)) {
  dir.create(output_folder, recursive = TRUE)
}
ggsave(file.path(output_folder, "gender_distribution_by_quarter_with_patterns.png"), gender_chart, width = 12, height = 16, dpi = 300)
ggsave(file.path(output_folder, "nationality_distribution_by_quarter_with_patterns.png"), nationality_chart, width = 12, height = 16, dpi = 300)
ggsave(file.path(output_folder, "origin_distribution_by_quarter_with_patterns.png"), origin_chart, width = 12, height = 16, dpi = 300)

cat("Stacked bar charts with patterns saved successfully in:", output_folder)

