
test_that("File saves as ODS", {
  ##Make note of files at start and at end
  start_all_files <- list.files("testfiles")

  convert_to_ods("testfiles/test_file.xlsx")

  end_all_files <- list.files("testfiles")

  expect_equal(sum(grepl("test_file.ods", start_all_files, fixed = TRUE)), 0)
  expect_equal(sum(grepl("test_file.ods", end_all_files, fixed = TRUE)), 1)

})

test_that("Returns error when file doesnt exist", {

  expect_error(convert_to_ods("abc.xlsx"))

})


test_that("File saves as XLSX", {
  ##Make note of files at start and at end
  start_all_files <- list.files("testfiles")

  convert_to_xlsx("testfiles/test_file1.ods")

  end_all_files <- list.files("testfiles")

  expect_equal(sum(grepl("test_file1.xlsx", start_all_files, fixed = TRUE)), 0)
  expect_equal(sum(grepl("test_file1.xlsx", end_all_files, fixed = TRUE)), 1)

})

test_that("Returns error when file doesnt exist", {

  expect_error(convert_to_xlsx("abc.ods"))

})


#Remove test file
 unlink("testfiles/test_file.ods")
 unlink("testfiles/test_file1.xlsx")

