# Libraries
library(dplyr)
library(stringr)
library(lubridate)

# Constants
COLS_OF_INTEREST <- c(
  "End.Customer.Account.Name",
  "Description",
  "Quantity",
  "Cost",
  "Commitment.Term",
  "Billing.Period",
  "Start.Period",
  "End.Period"
)

# Functions
# =========

get_changed <- function(old, new) {
  cols <- c(
    "Quantity",
    "Commitment.Term",
    "Billing.Period",
    "Cost"
  )
  
  old$compare <- old |> 
    select(all_of(cols)) |> 
    apply(1, paste0, collapse = "-") |> 
    str_squish()
  
  new$compare <- new |> 
    select(all_of(cols)) |> 
    apply(1, paste0, collapse = "-") |> 
    str_squish() 
  
  changed_subID <- full_join(old, new, by = "Subscription.ID", relationship = "many-to-many") |> 
    filter(compare.x != compare.y) |> 
    pull(Subscription.ID) |> 
    unique()

  return(changed_subID)
}

# Get subscriptions about to expire
all_new_data <- function(df, extra) {
  extra <- extra |> 
    select(all_of(c("Subscription.ID", "Renewal.date", "Expiration.date")))
  
  df <- df |> 
    select(all_of(c("Subscription.ID", COLS_OF_INTEREST))) |> 
    left_join(extra, by = "Subscription.ID") |> 
    arrange(End.Customer.Account.Name, Subscription.ID, Expiration.date)
  
  return(df)
}

get_near_expire <- function(df) {
    filter(df, as.Date(Expiration.date, format = "%d/%m/%Y") <= ceiling_date(Sys.Date() + months(1), unit = "month")) |> 
    arrange(desc(Commitment.Term), End.Customer.Account.Name, Subscription.ID, Expiration.date) |> 
    select(-Subscription.ID)
}

easy_compare_data <- function(old, new, sub) {
  old$Compare.Type <- "previous"
  new$Compare.Type <- "current"
  old_subID <- unique(old$Subscription.ID)
  new_subID <- unique(new$Subscription.ID)
  
  new_additions <- new |> 
    filter(!Subscription.ID %in% old_subID) |> 
    select(all_of(COLS_OF_INTEREST)) |> 
    arrange(End.Customer.Account.Name)
  
  gone <- old |> 
    filter(!Subscription.ID %in% new_subID) |> 
    select(all_of(COLS_OF_INTEREST)) |> 
    arrange(End.Customer.Account.Name)
  
  changed_subID <- get_changed(
    old = filter(old, !Subscription.ID %in% gone$Subscription.ID),
    new = filter(new, !Subscription.ID %in% new_additions$Subscription.ID)
  ) 
  
  changed <- rbind(old, new) |> 
    filter(Subscription.ID %in% changed_subID) |> 
    arrange(End.Customer.Account.Name, Subscription.ID, Compare.Type, Start.Period) |> 
    select(all_of(c(COLS_OF_INTEREST, "Compare.Type")))
  
  all <- all_new_data(new, sub)
  
  expire <- get_near_expire(all)
  
  return(list("Gone" = gone, 
              'New Additions' = new_additions, 
              "Changed" = changed, 
              "Expires soon" = expire,
              "View all" = all))
}


# Examples
#old <- read.csv("./Data/Invoice_Compare_Data Dec 2025.csv")
#new <- read.csv("./Data/Invoice_Compare_Data Jan 2026.csv")
#sub_data <- read.csv("./Data/Subscriptions_20260212_093956.csv")

#compare_data <- easy_compare_data(old, new)

