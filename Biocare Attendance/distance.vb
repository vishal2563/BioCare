Public Class distance

    Private Sub distance_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        'Calulating the distance in miles/kilo between two know points (Latitude/Longitude)
        'The Haversine formula according to Dr. Math.
        '    http://mathforum.org/library/drmath/view/51879.html
        'dlon = lon2 - lon1
        'dlat = lat2 - lat1
        'a = (sin(dlat / 2)) ^ 2 + cos(lat1) * cos(lat2) * (sin(dlon / 2)) ^ 2
        'c = 2 * atan2(sqrt(a), sqrt(1 - a))
        'd = R * c
        'Where()
        '     * dlon is the change in longitude
        '     * dlat is the change in latitude
        '     * c is the great circle distance in Radians.
        '     * R is the radius of a spherical Earth.
        '     * The locations of the two points in 
        '            spherical coordinates (longitude and 
        '            latitude) are lon1,lat1 and lon2, lat2.

        Dim dDistance As Double
        Dim dLat1InRad As Double
        Dim dLong1InRad As Double
        Dim dLat2InRad As Double
        Dim dLong2InRad As Double
        Dim dLongitude As Double
        Dim dLatitude As Double
        Dim a As Double
        Dim c As Double
        Dim EarthRadiusMiles As Integer


        'convert to Radians
        dLat1InRad = txtLocLAT1.Text
        dLat1InRad = dLat1InRad * (Math.PI / 180.0)

        dLong1InRad = txtLocLON1.Text
        dLong1InRad = dLong1InRad * (Math.PI / 180.0)

        dLat2InRad = txtLocLAT2.Text
        dLat2InRad = dLat2InRad * (Math.PI / 180.0)

        dLong2InRad = txtLocLON2.Text
        dLong2InRad = dLong2InRad * (Math.PI / 180.0)

        dLongitude = dLong2InRad - dLong1InRad
        dLatitude = dLat2InRad - dLat1InRad

        'Intermediate result a.
        a = Math.Pow(Math.Sin(dLatitude / 2.0), 2.0) + Math.Cos(dLat1InRad) * Math.Cos(dLat2InRad) * Math.Pow(Math.Sin(dLongitude / 2.0), 2.0)

        'Intermediate result c (great circle distance in Radians).
        c = 2.0 * Math.Asin(Math.Sqrt(a))

        'Distance.
        EarthRadiusMiles = 3956.0
        'const Double EarthRadiusKms = 6376.5;
        dDistance = EarthRadiusMiles * c

        txtDistance.Text = Format(dDistance, "#,##0.00")
    End Sub
End Class