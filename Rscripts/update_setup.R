
# update setup

update_setup  <- function(update_status = 0) {
    #move previous version of my designer
    if (update_status == 0){
        file.copy(from = "./src/bin/setup_aky.xlsb",
                to = "./Rscripts/", overwrite = TRUE)
        # move back and overwrite
        file.copy(from = "./Rscripts/setup_aky.xlsb",
                to = "./src/bin/setup_dev.xlsb", overwrite = TRUE)
    }
    # update the stable version if needed
    if (update_status == 1) {
       # previous stable version
         file.copy(from = "./setup.xlsb",
             to = "./Rscripts/setup_prev.xlsb", overwrite = TRUE)
       # update the new stable version
         file.copy(from = "./Rscripts/setup_aky.xlsb",
             to = "./setup.xlsb", overwrite = TRUE)
    }

    # revert back previous stable designer due to corrupt files.
    if (update_status == 2) {
         file.copy(from = "./Rscripts/setup_prev.xlsb",
             to = "./setup.xlsb", overwrite = TRUE)
         file.copy(from = "./Rscripts/setup_prev.xlsb",
                 to = "./src/bin/setup_dev.xlsb", overwrite = TRUE)
         file.copy(from = "./Rscripts/setup_prev.xlsb",
                 to = "./setup_aky.xlsb", overwrite = TRUE)
    }
    # just copy the mock file
    if (update_status == 3) {
        file.copy(from = "./src/.mock/setup_mock.xlsb",
                 to = "./src/bin/setup_aky.xlsb", overwrite = TRUE)
    }
}

#update the dev file in bin
update_setup(update_status = 0) #nolint
#update the file on root setup
update_setup(update_status = 1) #nolint
#replace the _aky setup by the _mock setup
update_setup(update_status = 3) #nolint
