test_that("Creates expected vbs filepaths", {

  expect_identical(vbs_file_path("test"),

  paste0('"', system.file("vbs", package = "odsconvertr"), "/test", '"')
  )

  expect_identical(vbs_file_path("empty_file123"),

                   paste0('"', system.file("vbs", package = "odsconvertr"), "/empty_file123", '"')
  )

})

