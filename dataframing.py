import pandas as p


def csv_load(filepath):
    csv = p.read_csv(filepath)
    return csv


# args provided by the buttons in gui should just be able to 'read' them to provide the info
def filter_csv(original_csv,desired_campaign):
    filtered_csv = original_csv.loc[(original_csv['campaign'] == desired_campaign),
            ["client","campaign","step","version","sent", "delivered", "opened", "opened_rate", "responded", "responded_rate", "interested_yes", "bounced",
                "bounce_rate", "opt_out", "opt_out_rate"]]

    return filtered_csv
