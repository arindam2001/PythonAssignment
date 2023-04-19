import pandas as pd

# Read the input sheets into pandas dataframes
df_user_ids = pd.read_excel("filename.xlsx", sheet_name="(Input) User IDs")
df_rigorbuilder_raw = pd.read_excel("filename.xlsx", sheet_name="(Input) Rigorbuilder RAW")

# Merge the dataframes on the "User ID" and "uid" columns
df_merged = pd.merge(df_user_ids, df_rigorbuilder_raw, left_on="User ID", right_on="uid")

# Create the "Team Wise Leaderboard" table
df_team_wise = df_merged.groupby("Team Name").agg({"total_statements": "mean", "total_reasons": "mean"}).reset_index()
df_team_wise.columns = ["Thinking Teams Leaderboard", "Average Statements", "Average Reasons"]




df_team_wise = df_team_wise.sort_values(by=["Average Statements", "Average Reasons"], ascending=False).reset_index(drop=True)
df_team_wise["Team Rank"] = df_team_wise.index + 1
df_team_wise = df_team_wise[["Team Rank", "Thinking Teams Leaderboard", "Average Statements", "Average Reasons"]]

# Create the "Leaderboard TeamWise" table
df_leaderboard = df_merged[["Name", "User ID", "total_statements", "total_reasons"]]
df_leaderboard.columns = ["Name", "UID", "No. of Statements", "No. of Reasons"]
df_leaderboard = df_leaderboard.sort_values(by=["No. of Statements", "No. of Reasons"], ascending=False).reset_index(drop=True)
df_leaderboard["Rank"] = df_leaderboard.index + 1
df_leaderboard = df_leaderboard[["Rank", "Name", "UID", "No. of Statements", "No. of Reasons"]]

# Write the output sheets with tables to the same Excel file
with pd.ExcelWriter("filename.xlsx", mode="a") as writer:
    df_team_wise.to_excel(writer, sheet_name="Leaderboard Individual  (Output)", index=False)
    df_leaderboard.to_excel(writer, sheet_name="Leaderboard TeamWise (Output)", index=False)




