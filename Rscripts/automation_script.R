#
source("./Rscripts/automate_class_creation.R")

update_setup(update_stable = 1)

# update by adding the stable
#update_setup(update_stable = 1)

# get the list of all the files in the class folder (for updating)
sink("./src/classes_list.txt")
ls  <- list.files("./src/classes")
ls  <- stringr::str_replace(ls, "\\.cls$", "\n")
cat(glue::glue("{ls}\n"))
sink()
