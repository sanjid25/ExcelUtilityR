test_that("Round Up Works", {
  # regular round of 2.5 rounds down
  expect_equal(round(2.5), 2)
  # round up of 2.5 rounds up
  expect_equal(roundUp(2.5), 3)
})
