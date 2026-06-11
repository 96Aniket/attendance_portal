from Execute.executesql import get_connection
from datetime import datetime, date

# =====================================================
# LOCATIONS
# =====================================================

def get_all_locations():
    try:
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT DISTINCT BrcName
            FROM tbl_ATTENDANCE_DATA
            WHERE BrcName IS NOT NULL
            ORDER BY BrcName
        """)

        rows = cursor.fetchall()
        cursor.close()
        conn.close()

        return [r[0] for r in rows]

    except Exception as e:
        print("get_all_locations error:", e)
        return []


# =====================================================
# MONTHS BY LOCATION
# =====================================================

def get_months_by_location(location):
    try:
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT DISTINCT
                FORMAT(PDate, 'yyyy-MM') AS MonthValue,
                FORMAT(PDate, 'MMM yyyy') AS MonthLabel
            FROM tbl_ATTENDANCE_DATA
            WHERE BrcName = ?
            ORDER BY MonthValue DESC
        """, (location,))

        rows = cursor.fetchall()
        cursor.close()
        conn.close()

        return [{"value": r[0], "label": r[1]} for r in rows]

    except Exception as e:
        print("get_months_by_location error:", e)
        return []


# =====================================================
# ORGANIZATIONS BY LOCATION + MONTH
# =====================================================

def get_orgs(location, month_year):
    try:
        year  = int(month_year.split("-")[0])
        month = int(month_year.split("-")[1])

        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT DISTINCT OrgName
            FROM tbl_ATTENDANCE_DATA
            WHERE BrcName = ?
              AND YEAR(PDate)  = ?
              AND MONTH(PDate) = ?
              AND ISNULL(delete_flag, 0) = 0
            ORDER BY OrgName
        """, (location, year, month))

        rows = cursor.fetchall()
        cursor.close()
        conn.close()

        return [r[0] for r in rows]

    except Exception as e:
        print("get_orgs error:", e)
        return []


# =====================================================
# ATTENDANCE DATA
# =====================================================

def fetch_attendance_data(year, month, location, organization):
    try:
        conn = get_connection()
        cursor = conn.cursor()

        query = """
            SELECT
                UserID,
                PDate,
                FullName,
                DsgName,
                BrcName,
                OrgName,
                AttendanceValue,
                Overtime
            FROM tbl_ATTENDANCE_DATA
            WHERE YEAR(PDate)  = ?
              AND MONTH(PDate) = ?
              AND ISNULL(delete_flag, 0) = 0
        """

        params = [year, month]

        if location:
            query += " AND BrcName = ?"
            params.append(location)

        if organization:
            query += " AND OrgName = ?"
            params.append(organization)

        query += " ORDER BY FullName, PDate"

        cursor.execute(query, params)
        rows = cursor.fetchall()

        result = []
        for r in rows:
            result.append({
                "UserID":          r[0],
                "PDate":           r[1],
                "FullName":        r[2],
                "DsgName":         r[3],
                "BrcName":         r[4],
                "OrgName":         r[5],
                "AttendanceValue": r[6],
                "Overtime":        r[7],
            })

        cursor.close()
        conn.close()

        return True, result

    except Exception as e:
        print("fetch_attendance_data error:", e)
        return False, str(e)
    