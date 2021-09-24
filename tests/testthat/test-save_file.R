
test_that("File saves as ODS", {
  ##Make note of files at start and at end
  start_all_files <- list.files("testfiles")

  convert_to_ods("testfiles/test_file.xlsx")

  end_all_files <- list.files("testfiles")

  expect_equal(sum(grepl("ods", start_all_files)), 0)
  expect_equal(sum(grepl("ods", end_all_files)), 1)

})

#Remove test file
unlink("testfiles/test_file.ods")


test_that("Absolute filepath argument works", {
  ##Make note of files at start and at end
  start_all_files <- list.files(system.file("testfiles", package = "odsconvertr"))

  convert_to_ods(paste0(system.file("testfiles", package = "odsconvertr"), "/test_file.xlsx"), relative_file_path = FALSE)

  end_all_files <- list.files(system.file("testfiles", package = "odsconvertr"))

  expect_equal(sum(grepl("ods", start_all_files)), 0)
  expect_equal(sum(grepl("ods", end_all_files)), 1)

})

#Remove test file
unlink(paste0(system.file("testfiles", package = "odsconvertr"), "/test_file.ods"))
