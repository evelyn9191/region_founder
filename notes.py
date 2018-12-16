# This is a part of code I may later use when solving certain problems.
# Not working, not correct, just notes.

for row_number in all_rows:
    # Verze z noci 11.12. 23.46 nez jsem ji zmenila na to, co tam je ted
    for row_number in all_rows:
        if df[df[user_data["cities_column"]].str.match(df_regions[row_number, "Okres"])]:
            get_region = df[user_data["regions_column"]] = df.regions[row_number, "Kraj"].copy()
        ws.cell(row=row_number + 2, column=user_data["regions_column"]).value = get_region


    if df[df[row_number, user_data["cities_column"]].str.match(df_regions[
                                                                   row_number, 0])]:  # TODO: KeyError: 'D' / ale vyhazuje to i kdyz tam dam integer, takze to bude spise o formatu toho vystupu
        df.loc[df["regions_column"]] = df_regions["B"]

        df["regions_column"] = df_regions["B"]


            found_region = "{}".format(df.regions[row_number, "Kraj"])
            ws.cell(row=row_number + 1, column=user_data["regions_column"]).value = found_region
    wb.save(output_file)
    print("Regions successfully matched to cities in", output_file)

if df[df[cities_column_number].str.contains(df_regions["Okres"])]:  # TODO: will have to use str.match or
# .apply instead