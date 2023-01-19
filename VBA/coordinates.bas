Attribute VB_Name = "Coordinates"
Function DISTANCE(latitude_A, longitude_A, latitude_B, longitude_B)
    ' Calculates the distance (in kilometers) from
    ' coordinate A to coordinate B
    ' VBA version of
    ' https://www.geeksforgeeks.org/program-distance-two-points-earth/
    ' coordinates have to be given as decimal numbers
    lon1 = Application.Radians(latitude_A)
    lon2 = Application.Radians(longitude_A)
    lat1 = Application.Radians(latitude_B)
    lat2 = Application.Radians(longitude_B)
      
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = Sin(dlat / 2) ^ 2 + Cos(lat1) * Cos(lat2) * Sin(dlon / 2) ^ 2

    DISTANCE = Round(2 * WorksheetFunction.Asin(Sqr(a)) * 6371, 2)
End Function
