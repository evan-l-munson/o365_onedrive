
# Libraries ---------------------------------------------------------------

library(Microsoft365R)
library(xlsx)

# Explore -----------------------------------------------------------------

# Connect to OneDrive
elm_od <- Microsoft365R::get_personal_onedrive()

# List files and folders
elm_od$list_files("Projects")
elm_od$list_files("Documents")
elm_od$list_files("Documents/Munson Stuff")

# Download file
elm_od$download_file("Documents/Munson Stuff/Stuff.xlsx")

# Read in file
stuff <- readxl::read_xlsx(path = "Stuff.xlsx")

# create toy dataset
toy <- sample(x = 1:5, size = 10, replace = TRUE)
toy_df <- as.data.frame(toy)

xlsx::write.xlsx(x = toy_df, 
                 file = "toy_df.xlsx")

# Save toy_df.xlsx back to OneDrive
elm_od$upload_file(src = "toy_df.xlsx", 
                   dest = "Projects/toy_df.xlsx")
