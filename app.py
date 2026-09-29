import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="HR Attendance & EDP",
    page_icon="https://img.icons8.com/scribby/50/f.png",
    layout="wide"
)


# ============================================================
# EMPLOYEE YANG DI-EXCLUDE
# ============================================================

EXCLUDED_EMPLOYEE_IDS = {
    "X25011047",   # Verdi Daivando
    "E1901446",    # Miftakhul Khoir
    "E23011010",   # Irawan Delfi Nurcahyo
    "E1701211",    # Syaiful Rahmat
    "E2001599",    # Tjakradidjaja Handany
    "E2301955",    # Dwi Ariawan Rahardjo
    "E2101731",    # R. E Wisnu Wardhana
    "E24011147",   # Fernandi Mahendrasusila
    "E1501052",    # Chandra Thomas
    "E24011146",   # Monika Ferolina Siallagan
    "E1601127",    # Ari Susanti
    "E24011093",   # Sabirin Muhtar
    "E1701180",    # Dea Rendra Kirana
    "E1301001",    # Ignatio Immanuel
    "E24011136",   # Hendra Gunawan
    "E2301956",    #Agus Gunawan
    "E1401003",    #Sugiharto Darmakusuma
    "E1401005",    #Deddy Manto
    "E1501064",    #Danny Aryanto
    "E1501069",    #Rizky Nur Arvianti
    "E1501071",    #Marcelus Ardiwinata
    "E2201818"     #Bambang Widjanarko E S
}


# ============================================================
# EMPLOYEE BLANK DIVISION -> LEGAL
# ============================================================

LEGAL_EMPLOYEE_IDS = {
    "E2101777",    # Arga Giantara Ikbar
    "E2201945",    # Amanda Putri Larasati
    "E2101776",    # Benny Suhancolo, SH
    "E2201933",    # Beta Duadja Sasana
    "E24011189",   # Bintang Ramadhan Prasetia Adriadi
    "E1701252",    # Eduard Kananites Aer
    "E25011331",   # Elisabet Oktalia Salsa Prajasari
    "E24011111",   # Firman
    "E2201895",    # Gulardi Achmad
    "E1701228",    # Rezki Meirza
    "E26011444",   # Syarifah Faizah
    "E25011321",   # Yugo Yehezkiel Wattimena
    "E24011148"    # Zeqwelin Ebestina Diana Ayu
}


# ============================================================
# EMPLOYEE BLANK DIVISION -> MARCOM
# ============================================================

MARCOM_EMPLOYEE_IDS = {
    "E2001607",    # Nick Steven
    "E2001620",    # Pandu Hadi Kesuma
    "E2301987",    # Ratih Sonia Setyawati. P
    "E26011461"    # Safira Salsabila
}


# ============================================================
# TITLE
# ============================================================

st.title("📊 HR Attendance & EDP Analytics")

st.write(
    "Upload file Attendance dan EDP untuk melakukan "
    "perhitungan otomatis."
)


# ============================================================
# UPLOAD FILE
# ============================================================

col1, col2 = st.columns(2)

with col1:
    attendance_file = st.file_uploader(
        "Upload Attendance Excel",
        type=["xlsx"],
        key="attendance"
    )

with col2:
    edp_file = st.file_uploader(
        "Upload EDP Excel",
        type=["xlsx"],
        key="edp"
    )


# ============================================================
# PROCESS
# ============================================================

if attendance_file is not None and edp_file is not None:

    try:

        # ====================================================
        # 1. READ ATTENDANCE
        # ====================================================

        attendance = pd.read_excel(
            attendance_file,
            keep_default_na=False
        )

        required_attendance_columns = [
            "Date",
            "Employee No",
            "Employee Name",
            "Division",
            "Attendance Status",
            "Time In",
            "Time Out",
            "Late In"
        ]

        missing_columns = [
            col
            for col in required_attendance_columns
            if col not in attendance.columns
        ]

        if missing_columns:
            st.error(
                "Kolom berikut tidak ditemukan di file Attendance:"
            )
            st.write(missing_columns)
            st.stop()

        attendance_work = attendance[
            required_attendance_columns
        ].copy()


        # ====================================================
        # 2. CLEAN DATA ATTENDANCE
        # ====================================================

        attendance_work["Employee No"] = (
            attendance_work["Employee No"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        attendance_work["Employee Name"] = (
            attendance_work["Employee Name"]
            .fillna("")
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

        attendance_work["Division"] = (
            attendance_work["Division"]
            .fillna("")
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

        attendance_work["Attendance Status"] = (
            attendance_work["Attendance Status"]
            .fillna("")
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )


        # ====================================================
        # 3. NORMALISASI DIVISION
        # ====================================================
        #
        # Tujuan:
        # - Blank / kosong tetap ""
        # - "N/A", "n/a", " N/A " menjadi "N/A"
        # - Division lain tetap seperti aslinya
        #
        # Jadi N/A TIDAK akan dianggap sebagai blank.
        # ====================================================

        attendance_work.loc[
            attendance_work["Division"].str.upper().eq("N/A"),
            "Division"
        ] = "N/A"


        # ====================================================
        # 4. FORMAT DATE
        # ====================================================

        attendance_work["Date"] = (
            pd.to_datetime(
                attendance_work["Date"],
                errors="coerce"
            )
            .dt.normalize()
        )


        # ====================================================
        # 5. HAPUS DATA TANGGAL / EMPLOYEE INVALID
        # ====================================================

        attendance_work = attendance_work[
            attendance_work["Date"].notna()
            &
            attendance_work["Employee No"].ne("")
        ].copy()


        # ====================================================
        # 6. ASSIGN BLANK DIVISION
        # ====================================================
        #
        # HANYA Division yang benar-benar kosong yang diproses.
        #
        # Blank + Employee Legal  -> Legal
        # Blank + Employee Marcom -> Marcom
        # Blank lainnya           -> Not Listed
        #
        # N/A TIDAK masuk sini.
        # ====================================================

        blank_division = (
            attendance_work["Division"].eq("")
        )


        # ----------------------------------------------------
        # Blank Division -> Legal
        # ----------------------------------------------------

        attendance_work.loc[
            blank_division
            &
            attendance_work["Employee No"].isin(
                LEGAL_EMPLOYEE_IDS
            ),
            "Division"
        ] = "Legal"


        # ----------------------------------------------------
        # Blank Division -> Marcom
        # ----------------------------------------------------

        attendance_work.loc[
            blank_division
            &
            attendance_work["Employee No"].isin(
                MARCOM_EMPLOYEE_IDS
            ),
            "Division"
        ] = "Marcom"


        # ----------------------------------------------------
        # Blank Division lainnya -> Not Listed
        # ----------------------------------------------------

        attendance_work.loc[
            attendance_work["Division"].eq(""),
            "Division"
        ] = "Not Listed"


        # ====================================================
        # 7. HINDARI DUPLICATE EMPLOYEE + DATE
        # ====================================================
        #
        # Satu employee pada satu tanggal hanya dihitung 1 hari.
        # ====================================================

        attendance_work = (
            attendance_work
            .drop_duplicates(
                subset=[
                    "Employee No",
                    "Date"
                ]
            )
        )


        # ====================================================
        # 8. COUNT DAY
        # ====================================================

        pattern = (
            "Present|Business Trip|Present Off|"
            "Absent|Leave|Permit"
        )

        attendance_work["Count Day"] = (
            attendance_work[
                "Attendance Status"
            ]
            .str.contains(
                pattern,
                case=False,
                regex=True,
                na=False
            )
            .astype(int)
        )


        # ====================================================
        # 9. CLOCK IN
        # ====================================================

        attendance_work["Clock In"] = (
            attendance_work["Time In"].ne("")
            |
            attendance_work[
                "Attendance Status"
            ]
            .str.contains(
                "Leave",
                case=False,
                na=False
            )
        ).astype(int)


        # ====================================================
        # 10. CLOCK OUT
        # ====================================================

        attendance_work["Clock Out"] = (
            attendance_work["Time Out"].ne("")
            |
            attendance_work[
                "Attendance Status"
            ]
            .str.contains(
                "Leave",
                case=False,
                na=False
            )
        ).astype(int)


        # ====================================================
        # 11. LATE
        # ====================================================

        attendance_work["Late In"] = pd.to_numeric(
            attendance_work["Late In"],
            errors="coerce"
        ).fillna(0)

        attendance_work["Late"] = (
            attendance_work["Late In"] > 0
        ).astype(int)


        # ====================================================
        # 12. LEAVE
        # ====================================================

        attendance_work["Leave"] = (
            attendance_work[
                "Attendance Status"
            ]
            .str.strip()
            .str.lower()
            .eq("leave")
        ).astype(int)


        # ====================================================
        # 13. SUMMARY ATTENDANCE PER EMPLOYEE
        # ====================================================

        summary = (
            attendance_work
            .groupby(
                [
                    "Employee No",
                    "Employee Name"
                ],
                as_index=False
            )
            .agg(
                Date=("Date", "count"),
                Division=("Division", "first"),
                Count_Day=("Count Day", "sum"),
                Clock_In=("Clock In", "sum"),
                Clock_Out=("Clock Out", "sum"),
                Late_Count=("Late", "sum"),
                Leave_Count=("Leave", "sum")
            )
        )


        # ====================================================
        # 14. CEK COUNT DAY
        # ====================================================

        summary["Count_Day"] = (
            pd.to_numeric(
                summary["Count_Day"],
                errors="coerce"
            )
            .fillna(0)
        )


        # ====================================================
        # 15. PERCENTAGE C.IN
        # ====================================================

        summary["C.In_Numeric"] = np.where(
            summary["Count_Day"] > 0,
            summary["Clock_In"] /
            summary["Count_Day"],
            0
        )

        summary["C.In_Numeric"] = (
            summary["C.In_Numeric"]
            .clip(0, 1)
        )


        # ====================================================
        # 16. PERCENTAGE C.OUT
        # ====================================================

        summary["C.Out_Numeric"] = np.where(
            summary["Count_Day"] > 0,
            summary["Clock_Out"] /
            summary["Count_Day"],
            0
        )

        summary["C.Out_Numeric"] = (
            summary["C.Out_Numeric"]
            .clip(0, 1)
        )


        # ====================================================
        # 17. PERCENTAGE LATE
        # ====================================================

        summary["Late_Numeric"] = np.where(
            summary["Count_Day"] > 0,
            summary["Late_Count"] /
            summary["Count_Day"],
            0
        )

        summary["Late_Numeric"] = (
            summary["Late_Numeric"]
            .clip(0, 1)
        )


        # ====================================================
        # 18. LEAVE PERCENTAGE
        # ====================================================

        summary["Leave_EDP_Numeric"] = np.where(
            summary["Count_Day"] > 0,
            summary["Leave_Count"] /
            summary["Count_Day"],
            0
        )

        summary["Leave_EDP_Numeric"] = (
            summary["Leave_EDP_Numeric"]
            .clip(0, 1)
        )


        # ====================================================
        # 19. FORMAT DISPLAY ATTENDANCE
        # ====================================================

        summary["C.In"] = (
            summary["C.In_Numeric"]
            .map("{:.0%}".format)
        )

        summary["C.Out"] = (
            summary["C.Out_Numeric"]
            .map("{:.0%}".format)
        )

        summary["Late"] = (
            summary["Late_Numeric"]
            .map("{:.0%}".format)
        )


        # ====================================================
        # 20. SORT EMPLOYEE
        # ====================================================

        summary = (
            summary
            .sort_values("Employee Name")
            .reset_index(drop=True)
        )


        # ====================================================
        # 21. READ EDP
        # ====================================================

        edp = pd.read_excel(
            edp_file,
            keep_default_na=False
        )

        required_edp_columns = [
            "Employee Name",
            "Date"
        ]

        missing_edp_columns = [
            col
            for col in required_edp_columns
            if col not in edp.columns
        ]

        if missing_edp_columns:
            st.error(
                "Kolom berikut tidak ditemukan di file EDP:"
            )
            st.write(missing_edp_columns)
            st.stop()


        # ====================================================
        # 22. CLEAN EMPLOYEE NAME EDP
        # ====================================================

        edp["Employee Name"] = (
            edp["Employee Name"]
            .fillna("")
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )


        # ====================================================
        # 23. AMBIL EMPLOYEE NO DARI EDP
        # ====================================================
        #
        # Contoh:
        #
        # (E2201880) Sapto Handoyo Jati
        #
        # menjadi:
        #
        # Employee No   = E2201880
        # Employee Name = Sapto Handoyo Jati
        # ====================================================

        edp["Employee No"] = (
            edp["Employee Name"]
            .str.extract(
                r"\((.*?)\)",
                expand=False
            )
            .fillna("")
            .str.strip()
        )


        # ====================================================
        # 24. BERSIHKAN NAMA EDP
        # ====================================================

        edp["Employee Name"] = (
            edp["Employee Name"]
            .str.replace(
                r"\(.*?\)",
                "",
                regex=True
            )
            .str.strip()
        )


        # ====================================================
        # 25. FORMAT DATE EDP
        # ====================================================

        edp["Date"] = (
            pd.to_datetime(
                edp["Date"],
                errors="coerce"
            )
            .dt.normalize()
        )


        # ====================================================
        # 26. HAPUS DATA INVALID EDP
        # ====================================================

        edp = edp[
            edp["Employee No"].ne("")
            &
            edp["Date"].notna()
        ].copy()


        # ====================================================
        # 27. ONE EMPLOYEE ONE DATE
        # ====================================================

        edp = (
            edp
            .drop_duplicates(
                subset=[
                    "Employee No",
                    "Date"
                ]
            )
        )


        # ====================================================
        # 28. EDP SUMMARY
        # ====================================================

        edp_summary = (
            edp
            .groupby(
                "Employee No",
                as_index=False
            )
            .agg(
                EDP_Count=(
                    "Date",
                    "nunique"
                )
            )
        )


        # ====================================================
        # 29. EDP ONLY
        # ====================================================

        edp_only = (
            summary[
                [
                    "Employee No",
                    "Employee Name",
                    "Count_Day"
                ]
            ]
            .merge(
                edp_summary,
                on="Employee No",
                how="left"
            )
        )

        edp_only["EDP_Count"] = (
            edp_only["EDP_Count"]
            .fillna(0)
        )

        edp_only["EDP_Numeric"] = np.where(
            edp_only["Count_Day"] > 0,
            edp_only["EDP_Count"]
            /
            edp_only["Count_Day"],
            0
        )

        edp_only["EDP_Numeric"] = (
            edp_only["EDP_Numeric"]
            .clip(0, 1)
        )

        edp_only["EDP"] = (
            edp_only["EDP_Numeric"]
            .map("{:.0%}".format)
        )


        # ====================================================
        # 30. MERGE EDP KE SUMMARY
        # ====================================================

        summary = summary.merge(
            edp_summary,
            on="Employee No",
            how="left"
        )


        # ====================================================
        # 31. EMPLOYEE TANPA EDP
        # ====================================================

        summary["EDP_Count"] = (
            summary["EDP_Count"]
            .fillna(0)
        )


        # ====================================================
        # 32. EDP TOTAL
        # ====================================================
        #
        # EDP file + Leave
        # ====================================================

        summary["EDP_Total"] = (
            summary["EDP_Count"]
            +
            summary["Leave_Count"]
        )


        # ====================================================
        # 33. BATASI EDP TOTAL
        # ====================================================

        summary["EDP_Total"] = np.minimum(
            summary["EDP_Total"],
            summary["Count_Day"]
        )


        # ====================================================
        # 34. EDP NUMERIC
        # ====================================================

        summary["EDP_Numeric"] = np.where(
            summary["Count_Day"] > 0,
            summary["EDP_Total"]
            /
            summary["Count_Day"],
            0
        )

        summary["EDP_Numeric"] = (
            summary["EDP_Numeric"]
            .clip(0, 1)
        )


        # ====================================================
        # 35. FORMAT EDP
        # ====================================================

        summary["EDP"] = (
            summary["EDP_Numeric"]
            .map("{:.0%}".format)
        )


        # ====================================================
        # 36. PISAHKAN EXCLUDED
        # ====================================================
        #
        # Excluded:
        # - TETAP ada di summary
        # - TETAP dihitung sampai EDP
        # - masuk sheet "Excluded"
        #
        # Tetapi:
        # - TIDAK masuk Employee Summary
        # - TIDAK masuk KPI
        # - TIDAK masuk Average Division
        # - TIDAK masuk Average Overall
        # ====================================================

        excluded_summary = (
            summary[
                summary["Employee No"].isin(
                    EXCLUDED_EMPLOYEE_IDS
                )
            ]
            .copy()
        )

        summary_active = (
            summary[
                ~summary["Employee No"].isin(
                    EXCLUDED_EMPLOYEE_IDS
                )
            ]
            .copy()
        )


        # ====================================================
        # 37. SUMMARY FINAL
        # ====================================================

        summary_final = (
            summary_active[
                [
                    "Employee No",
                    "Employee Name",
                    "Division",
                    "C.In",
                    "C.Out",
                    "Late",
                    "EDP"
                ]
            ]
            .sort_values("Employee Name")
            .reset_index(drop=True)
        )


        # ====================================================
        # SUCCESS
        # ====================================================

        st.success(
            "Perhitungan berhasil!"
        )


        # ====================================================
        # KPI
        # ====================================================

        total_employee = (
            summary_final["Employee No"]
            .nunique()
        )

        total_division = (
            summary_final["Division"]
            .nunique()
        )


        # ====================================================
        # AVERAGE PERFORMANCE PER DIVISION
        # ====================================================

        division_average = (
            summary_active
            .groupby("Division")
            .agg(
                Employee_Count=(
                    "Employee No",
                    "nunique"
                ),
                Average_C_In=(
                    "C.In_Numeric",
                    "mean"
                ),
                Average_C_Out=(
                    "C.Out_Numeric",
                    "mean"
                ),
                Average_Late=(
                    "Late_Numeric",
                    "mean"
                ),
                Average_EDP=(
                    "EDP_Numeric",
                    "mean"
                )
            )
            .reset_index()
        )


        # ====================================================
        # AVERAGE PERFORMANCE
        # ====================================================
        #
        # Average =
        # (Average C.In + Average C.Out + Average EDP) / 3
        #
        # Late tidak dimasukkan.
        # ====================================================

        division_average["Average"] = (
            division_average["Average_C_In"]
            +
            division_average["Average_C_Out"]
            +
            division_average["Average_EDP"]
        ) / 3

        division_average["Average"] = (
            division_average["Average"]
            .clip(0, 1)
        )


        # ====================================================
        # AVERAGE OVERALL
        # ====================================================

        if len(division_average) > 0:
            average_overall = (
                division_average["Average"]
                .mean()
            )
        else:
            average_overall = 0


        # ====================================================
        # KPI DISPLAY
        # ====================================================

        k1, k2, k3 = st.columns(3)

        k1.metric(
            "Total Employee",
            total_employee
        )

        k2.metric(
            "Total Division",
            total_division
        )

        k3.metric(
            "Average Overall",
            f"{average_overall:.0%}"
        )


        # ====================================================
        # AVERAGE PERFORMANCE PER DIVISION
        # ====================================================

        st.subheader(
            "📈 Average Performance per Division"
        )

        st.caption(
            "Average C.In, C.Out, Late, dan EDP dihitung "
            "dengan rata-rata persentase masing-masing employee "
            "dalam division, sama seperti fungsi AVERAGE di Excel."
        )


        # ====================================================
        # COPY DISPLAY
        # ====================================================

        division_average_display = (
            division_average.copy()
        )


        # ====================================================
        # FORMAT PERCENTAGE
        # ====================================================

        percentage_columns = [
            "Average_C_In",
            "Average_C_Out",
            "Average_Late",
            "Average_EDP",
            "Average"
        ]

        for col in percentage_columns:
            division_average_display[col] = (
                division_average_display[col]
                .map("{:.0%}".format)
            )


        # ====================================================
        # RENAME
        # ====================================================

        division_average_display = (
            division_average_display
            .rename(
                columns={
                    "Employee_Count": "Total Emp",
                    "Average_C_In": "Avg C.In",
                    "Average_C_Out": "Avg C.Out",
                    "Average_Late": "Avg Late",
                    "Average_EDP": "Avg EDP",
                    "Average": "Average"
                }
            )
            .sort_values("Division")
        )


        # ====================================================
        # DISPLAY
        # ====================================================

        st.dataframe(
            division_average_display,
            use_container_width=True,
            hide_index=True
        )


        # ====================================================
        # FILTER DIVISION
        # ====================================================

        st.subheader(
            "👥 Employee Summary"
        )

        divisions = sorted(
            summary_final["Division"]
            .dropna()
            .unique()
            .tolist()
        )

        selected_division = st.selectbox(
            "Pilih Division",
            ["All"] + divisions
        )


        if selected_division == "All":

            display_data = (
                summary_final.copy()
            )

        else:

            display_data = (
                summary_final[
                    summary_final["Division"]
                    ==
                    selected_division
                ]
                .copy()
            )


        # ====================================================
        # DISPLAY EMPLOYEE SUMMARY
        # ====================================================

        st.dataframe(
            display_data,
            use_container_width=True,
            hide_index=True
        )


        # ====================================================
        # VERIFICATION AVERAGE
        # ====================================================

        with st.expander(
            "🧮 Verification Average seperti Excel"
        ):

            st.write(
                "Perhitungan berikut menggunakan "
                "`mean()` pada nilai persentase employee."
            )


            if selected_division == "All":

                verification = (
                    summary_active[
                        [
                            "Employee No",
                            "Employee Name",
                            "Division",
                            "C.In_Numeric",
                            "C.Out_Numeric",
                            "Late_Numeric",
                            "EDP_Numeric"
                        ]
                    ]
                    .copy()
                )

            else:

                verification = (
                    summary_active[
                        summary_active["Division"]
                        ==
                        selected_division
                    ][
                        [
                            "Employee No",
                            "Employee Name",
                            "Division",
                            "C.In_Numeric",
                            "C.Out_Numeric",
                            "Late_Numeric",
                            "EDP_Numeric"
                        ]
                    ]
                    .copy()
                )


            verification_display = (
                verification.copy()
            )


            for col in [
                "C.In_Numeric",
                "C.Out_Numeric",
                "Late_Numeric",
                "EDP_Numeric"
            ]:

                verification_display[col] = (
                    verification_display[col]
                    .map("{:.2%}".format)
                )


            verification_display = (
                verification_display
                .rename(
                    columns={
                        "C.In_Numeric": "C.In",
                        "C.Out_Numeric": "C.Out",
                        "Late_Numeric": "Late",
                        "EDP_Numeric": "EDP"
                    }
                )
            )


            st.dataframe(
                verification_display,
                use_container_width=True,
                hide_index=True
            )


            # ----------------------------------------------
            # EXCEL-LIKE AVERAGE
            # ----------------------------------------------

            if len(verification) > 0:

                st.write(
                    "### Hasil AVERAGE"
                )

                av1, av2, av3, av4 = st.columns(4)

                av1.metric(
                    "AVERAGE C.In",
                    f"{verification['C.In_Numeric'].mean():.2%}"
                )

                av2.metric(
                    "AVERAGE C.Out",
                    f"{verification['C.Out_Numeric'].mean():.2%}"
                )

                av3.metric(
                    "AVERAGE Late",
                    f"{verification['Late_Numeric'].mean():.2%}"
                )

                av4.metric(
                    "AVERAGE EDP",
                    f"{verification['EDP_Numeric'].mean():.2%}"
                )

                st.info(
                    f"Jumlah employee yang di-average: "
                    f"{len(verification)} employee"
                )


        # ====================================================
        # DOWNLOAD EXCEL
        # ====================================================

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            # =================================================
            # SHEET AVERAGE DIVISION
            # =================================================

            division_average_display.to_excel(
                writer,
                sheet_name="Average Division",
                index=False
            )


            # =================================================
            # SHEET EXCLUDED
            # =================================================

            excluded_excel_columns = [
                "Employee No",
                "Employee Name",
                "Division",
                "C.In",
                "C.Out",
                "Late",
                "EDP"
            ]

            excluded_data = (
                excluded_summary[
                    excluded_excel_columns
                ]
                .sort_values("Employee Name")
                .reset_index(drop=True)
            )

            excluded_data.to_excel(
                writer,
                sheet_name="Excluded",
                index=False
            )

            worksheet = writer.sheets["Excluded"]

            worksheet.column_dimensions["A"].width = 18
            worksheet.column_dimensions["B"].width = 35
            worksheet.column_dimensions["C"].width = 20


            # =================================================
            # SHEET PER DIVISION
            # =================================================

            used_sheet_names = {
                "Average Division",
                "Excluded"
            }


            for division, data in (
                summary_final
                .groupby("Division")
            ):

                # ---------------------------------------------
                # CLEAN SHEET NAME
                # ---------------------------------------------

                sheet_name = str(
                    division
                ).strip()


                # ---------------------------------------------
                # INVALID EXCEL CHARACTERS
                # ---------------------------------------------

                sheet_name = re.sub(
                    r'[:\\/?*\[\]]',
                    "_",
                    sheet_name
                )


                # ---------------------------------------------
                # EMPTY NAME
                # ---------------------------------------------

                if not sheet_name:
                    sheet_name = "Not Listed"


                # ---------------------------------------------
                # MAX 31 CHARACTERS
                # ---------------------------------------------

                sheet_name = sheet_name[:31]


                # ---------------------------------------------
                # DUPLICATE SHEET NAME
                # ---------------------------------------------

                original_sheet_name = sheet_name
                counter = 1

                while sheet_name in used_sheet_names:

                    suffix = f"_{counter}"

                    sheet_name = (
                        original_sheet_name[
                            :31 - len(suffix)
                        ]
                        +
                        suffix
                    )

                    counter += 1

                used_sheet_names.add(
                    sheet_name
                )


                # ---------------------------------------------
                # SORT
                # ---------------------------------------------

                data = (
                    data
                    .sort_values(
                        "Employee Name"
                    )
                )


                # ---------------------------------------------
                # KOLOM EXCEL
                # ---------------------------------------------

                excel_columns = [
                    "Employee No",
                    "Employee Name",
                    "C.In",
                    "C.Out",
                    "Late",
                    "EDP"
                ]

                data_excel = (
                    data[
                        excel_columns
                    ]
                    .sort_values("Employee Name")
                    .reset_index(drop=True)
                )


                # ---------------------------------------------
                # WRITE
                # ---------------------------------------------

                data_excel.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

                worksheet = writer.sheets[sheet_name]

                worksheet.column_dimensions["A"].width = 18
                worksheet.column_dimensions["B"].width = 35


        output.seek(0)


        # ====================================================
        # DOWNLOAD
        # ====================================================

        st.download_button(
            label=(
                "📥 Download Excel "
                "Per Division"
            ),
            data=output,
            file_name="Summary_2026.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            )
        )


        # ====================================================
        # DEBUG DIVISION
        # ====================================================

        with st.expander(
            "🔎 Cek Division"
        ):

            st.write(
                "### Daftar Division setelah mapping"
            )

            division_check = (
                summary_active[
                    [
                        "Employee No",
                        "Employee Name",
                        "Division"
                    ]
                ]
                .sort_values(
                    [
                        "Division",
                        "Employee Name"
                    ]
                )
                .reset_index(drop=True)
            )

            st.dataframe(
                division_check,
                use_container_width=True,
                hide_index=True
            )

            st.write(
                "### Jumlah Employee per Division"
            )

            division_count_check = (
                summary_active
                .groupby("Division")
                .agg(
                    Employee_Count=(
                        "Employee No",
                        "nunique"
                    )
                )
                .reset_index()
                .sort_values("Division")
            )

            st.dataframe(
                division_count_check,
                use_container_width=True,
                hide_index=True
            )


        # ====================================================
        # DEBUG EDP
        # ====================================================

        with st.expander(
            "🔎 Lihat Perhitungan EDP"
        ):

            st.write(
                "### EDP dari file EDP saja"
            )

            st.dataframe(
                edp_only,
                use_container_width=True,
                hide_index=True
            )


            st.write(
                "### Hasil akhir EDP"
            )

            st.dataframe(
                summary_active[
                    [
                        "Employee No",
                        "Employee Name",
                        "Division",
                        "Count_Day",
                        "Leave_Count",
                        "EDP_Count",
                        "EDP_Total",
                        "EDP"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )


            st.write(
                "### Employee Excluded"
            )

            st.dataframe(
                excluded_summary[
                    [
                        "Employee No",
                        "Employee Name",
                        "Division",
                        "Count_Day",
                        "Leave_Count",
                        "EDP_Count",
                        "EDP_Total",
                        "EDP"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )


    # ========================================================
    # ERROR HANDLING
    # ========================================================

    except Exception as e:

        st.error(
            f"Terjadi error: {e}"
        )

        st.exception(e)