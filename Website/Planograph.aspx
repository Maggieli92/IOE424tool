<%@ page language="VB" MasterPageFile="MasterPage.master" %>
<script runat="server">

Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
 
        Dim session_name As Label
        session_name = Master.FindControl("Session_Name")
        session_name.Text = "Logged in as : " & UCase(Session("user"))
        session_name.Visible = True
    End Sub
    
    Protected Sub CreateCSV(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
              
        Dim PartNumber As String
        Dim PkgCode As String
        Dim PackageWidth As Double
        Dim PackageHeight As Double
        Dim PackageDepth As Double
        Dim UnitOfIssue As Double
        Dim QtyOnHand As Double
        Dim MaterialType As String
        Dim Weight As Double
        Dim Avgdemand As Double
        Dim MaterialNumber As String
        Dim MaterialBrand As String
        Dim PartClass As String
        Dim FreeStock As Double
        Dim SIL As Integer
        Dim OriginalSIL As Integer
        Dim PartPrefix As String
        Dim PartBase As String
        Dim PartSuffix As String
        Dim Zone As String
        Dim PlanoSite As String
        Dim Location As String
        Dim InventoryOnHand As Integer
        Dim MultSale As Double
        Dim Hazmat As String
        Dim PartVolume As Double
        Dim PartVolumeSIL As Double
        Dim PackagePerimeter As Double
        Dim MezzanineLocation As String
        Dim MezzanineSize As String
        Dim strExtraShelfSpace
        Dim strPreviousShelf
        Dim PkgDim1 As Double
        Dim PkgDim2 As Double
        Dim PkgDim3 As Double
        Dim ShelfDim1 As Double
        Dim ShelfDim2 As Double
        Dim ShelfDim3 As Double
        Dim SuperBulkBase As String
        Dim SuperDuperBulk As String
        
        Dim PalletMaxFit As Integer
        
        Dim Overstock As Double
        Dim OverstockVolume As Double
        Dim OverstockPalletsNeeded As Double
        Dim SuperBulkPalletVolume As Double
        
        Dim SizingQty As Integer
        
        
        
        Dim MaxFitsOrientation
        
        Dim strdim1switch As Double
        Dim strdim2switch As Double
        Dim strdim3switch As Double
        
        Dim OpeningHeight As Double
        Dim OpeningWidth As Double
        Dim OpeningDepth As Double
        Dim OpeningVolume As Double
        Dim strCSV As String
        Dim strCSV2 As String
        Dim strCSV3 As String
        strCSV3 = ""
        Dim MyArray(8, 0) As Object
        Dim MyArray2(8, 0) As Object
        Dim MyArray3(8, 0) As Object
        Dim MyArray4(8, 0) As Object
        Dim MyBinArray(1, 0) As Object
        Dim MyPartArray(21, 0) As Object
        Dim k As Integer
        k = 0
        Dim w As Integer
        w = 0
        Dim x As Integer
        x = 0
        Dim y As Integer
        y = 0
        Dim z As Integer
        z = 0
        Dim MyArrayMaxFits(6) As Object
        Dim MyArrayMaxFitsMezzanine(6) As Object
        Dim MaxFit As Decimal
        Dim LargestLocation As String
        
        Dim csvheaderstring As String
        csvheaderstring = " "
        
        Dim MezzanineShelfCounter As Integer
        MezzanineShelfCounter = 0
        Dim BulkShelfCounter As Integer
        BulkShelfCounter = 0
        Dim MouldingShelfCounter As Integer
        MouldingShelfCounter = 0
        Dim SuperBulkShelfCounter As Integer
        SuperBulkShelfCounter = 0
        
        Dim FirstMezzanineShelf As String
        Dim FirstBulkShelf As String
        Dim FirstMouldingShelf As String
        Dim FirstSuperBulkShelf As String
        
        Dim MyArrayDisplaySizing(22)

        'Build Mezzanine shelf size array
        Dim mySqlConnection As SqlConnection
        Dim mySqlCommand As SqlCommand
        Dim myReader As SqlDataReader
        mySqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString1").ConnectionString)
        mySqlCommand = New SqlCommand("SELECT * FROM LocationSizes Where LocationName='" & Mezzanine.SelectedValue & "' and WhseCategory='Mezzanine'", mySqlConnection)

       
        mySqlConnection.Open()
        myReader = mySqlCommand.ExecuteReader()
        If myReader.HasRows = True Then
            
            Do While myReader.Read()
                
                
                
                ShelfDim1 = myReader("OpeningWidth")
                ShelfDim2 = myReader("OpeningHeight")
                ShelfDim3 = myReader("OpeningDepth")
                
                If ShelfDim3 > ShelfDim2 Then
                    strdim3switch = ShelfDim3
                    strdim2switch = ShelfDim2
                    ShelfDim2 = strdim3switch
                    ShelfDim3 = strdim2switch
                End If

                If ShelfDim2 > ShelfDim1 Then
                    strdim2switch = ShelfDim2
                    strdim1switch = ShelfDim1
                    ShelfDim1 = strdim2switch
                    ShelfDim2 = strdim1switch
                End If
	
                If ShelfDim3 > ShelfDim2 Then
                    strdim3switch = ShelfDim3
                    strdim2switch = ShelfDim2
                    ShelfDim2 = strdim3switch
                    ShelfDim3 = strdim2switch
                End If
              
                ReDim Preserve MyArray(8, k)

                MyArray(0, k) = myReader("LocationSizeCode")
                MyArray(1, k) = myReader("OpeningWidth")
                MyArray(2, k) = myReader("OpeningHeight")
                MyArray(3, k) = myReader("OpeningDepth")
                MyArray(4, k) = myReader("WhseArea")
                MyArray(5, k) = myReader("OpeningWidth") * myReader("OpeningHeight") * myReader("OpeningDepth")
                MyArray(6, k) = ShelfDim1
                MyArray(7, k) = ShelfDim2
                MyArray(8, k) = ShelfDim3
                k = k + 1
                MezzanineShelfCounter = MezzanineShelfCounter + 1
            Loop
            
        End If
        mySqlConnection.Close()
        
        Dim MyArrayBinFinder(7, k - 1) As Object
       
        Dim i As Integer
        
        For i = 0 To k - 1
            Response.Write(MyArray(0, i) & " | " & MyArray(1, i) & " | " & MyArray(2, i) & " | " & MyArray(3, i) & " | " & MyArray(4, i) & " | Total Vol=" & MyArray(5, i) & "<br>")
        Next
        Response.Write("<br><br>")
        Dim j
        Dim m
        Dim s
        Dim p
        Dim t
        Dim b
        b = 0
        
        'Sort the array
        For p = k - 1 To 0 Step -1
            For j = k - 2 To 0 Step -1
                If MyArray(5, j) < MyArray(5, j + 1) Then
                    For m = 0 To 8
                        s = MyArray(m, j + 1)
                        MyArray(m, j + 1) = MyArray(m, j)
                        MyArray(m, j) = s
                    Next
                End If
            Next
        Next
        
         FirstMezzanineShelf = MyArray(0, 0)
        
        For i = k - 1 To 0 Step -1
            csvheaderstring = csvheaderstring & "," & (MyArray(0, i))
        Next
        
        
        
        'Build Bulk Area shelf size array
      
        mySqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString1").ConnectionString)
        mySqlCommand = New SqlCommand("SELECT * FROM LocationSizes Where LocationName='" & Mezzanine.SelectedValue & "' and WhseCategory='Bulk' ", mySqlConnection)

       
        mySqlConnection.Open()
        myReader = mySqlCommand.ExecuteReader()
        If myReader.HasRows = True Then
            
            Do While myReader.Read()
                
                ShelfDim1 = myReader("OpeningWidth")
                ShelfDim2 = myReader("OpeningHeight")
                ShelfDim3 = myReader("OpeningDepth")
                
                If ShelfDim3 > ShelfDim2 Then
                    strdim3switch = ShelfDim3
                    strdim2switch = ShelfDim2
                    ShelfDim2 = strdim3switch
                    ShelfDim3 = strdim2switch
                End If

                If ShelfDim2 > ShelfDim1 Then
                    strdim2switch = ShelfDim2
                    strdim1switch = ShelfDim1
                    ShelfDim1 = strdim2switch
                    ShelfDim2 = strdim1switch
                End If
	
                If ShelfDim3 > ShelfDim2 Then
                    strdim3switch = ShelfDim3
                    strdim2switch = ShelfDim2
                    ShelfDim2 = strdim3switch
                    ShelfDim3 = strdim2switch
                End If
               
                ReDim Preserve MyArray2(8, x)

                MyArray2(0, x) = myReader("LocationSizeCode")
                MyArray2(1, x) = myReader("OpeningWidth")
                MyArray2(2, x) = myReader("OpeningHeight")
                MyArray2(3, x) = myReader("OpeningDepth")
                MyArray2(4, x) = myReader("WhseArea")
                MyArray2(5, x) = myReader("OpeningWidth") * myReader("OpeningHeight") * myReader("OpeningDepth")
                MyArray2(6, x) = ShelfDim1
                MyArray2(7, x) = ShelfDim2
                MyArray2(8, x) = ShelfDim3
                x = x + 1
                BulkShelfCounter = BulkShelfCounter + 1
            Loop
            
        End If
        mySqlConnection.Close()
        
      b = 0
        
        'Sort the array
        For p = x - 1 To 0 Step -1
            For j = x - 2 To 0 Step -1
                If MyArray2(5, j) < MyArray2(5, j + 1) Then
                    For m = 0 To 8
                        s = MyArray2(m, j + 1)
                        MyArray2(m, j + 1) = MyArray2(m, j)
                        MyArray2(m, j) = s
                    Next
                End If
            Next
        Next
        
        FirstBulkShelf = MyArray2(0, 0)
        
        For i = x - 1 To 0 Step -1
            csvheaderstring = csvheaderstring & "," & (MyArray2(0, i))
        Next
       
        
        'Build Moulding Area shelf size array
      
        mySqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString1").ConnectionString)
        mySqlCommand = New SqlCommand("SELECT * FROM LocationSizes Where LocationName='" & Mezzanine.SelectedValue & "' and WhseCategory='Moulding' ", mySqlConnection)

       
        mySqlConnection.Open()
        myReader = mySqlCommand.ExecuteReader()
        If myReader.HasRows = True Then
            
            Do While myReader.Read()
                
                ShelfDim1 = myReader("OpeningWidth")
                ShelfDim2 = myReader("OpeningHeight")
                ShelfDim3 = myReader("OpeningDepth")
                
                If ShelfDim3 > ShelfDim2 Then
                    strdim3switch = ShelfDim3
                    strdim2switch = ShelfDim2
                    ShelfDim2 = strdim3switch
                    ShelfDim3 = strdim2switch
                End If

                If ShelfDim2 > ShelfDim1 Then
                    strdim2switch = ShelfDim2
                    strdim1switch = ShelfDim1
                    ShelfDim1 = strdim2switch
                    ShelfDim2 = strdim1switch
                End If
	
                If ShelfDim3 > ShelfDim2 Then
                    strdim3switch = ShelfDim3
                    strdim2switch = ShelfDim2
                    ShelfDim2 = strdim3switch
                    ShelfDim3 = strdim2switch
                End If
            
                ReDim Preserve MyArray4(8, z)

                MyArray4(0, z) = myReader("LocationSizeCode")
                MyArray4(1, z) = myReader("OpeningWidth")
                MyArray4(2, z) = myReader("OpeningHeight")
                MyArray4(3, z) = myReader("OpeningDepth")
                MyArray4(4, z) = myReader("WhseArea")
                MyArray4(5, z) = myReader("OpeningWidth") * myReader("OpeningHeight") * myReader("OpeningDepth")
                MyArray4(6, z) = ShelfDim1
                MyArray4(7, z) = ShelfDim2
                MyArray4(8, z) = ShelfDim3
                z = z + 1
                MouldingShelfCounter = MouldingShelfCounter + 1
                
            Loop
            
        End If
        mySqlConnection.Close()
        
        b = 0
        
        'Sort the array
        For p = z - 1 To 0 Step -1
            For j = z - 2 To 0 Step -1
                If MyArray4(5, j) < MyArray4(5, j + 1) Then
                    For m = 0 To 8
                        s = MyArray4(m, j + 1)
                        MyArray4(m, j + 1) = MyArray4(m, j)
                        MyArray4(m, j) = s
                    Next
                End If
            Next
        Next
        
        FirstMouldingShelf = MyArray4(0, 0)
        
        For i = z - 1 To 0 Step -1
            csvheaderstring = csvheaderstring & "," & (MyArray4(0, i))
        Next
        
        
        
        'Build Super Bulk Area shelf size array
      
        mySqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString1").ConnectionString)
        mySqlCommand = New SqlCommand("SELECT * FROM LocationSizes Where LocationName='" & Mezzanine.SelectedValue & "' and WhseCategory='Super Bulk' ", mySqlConnection)

       
        mySqlConnection.Open()
        myReader = mySqlCommand.ExecuteReader()
        If myReader.HasRows = True Then
            
            Do While myReader.Read()
                
                ShelfDim1 = myReader("OpeningWidth")
                ShelfDim2 = myReader("OpeningHeight")
                ShelfDim3 = myReader("OpeningDepth")
                
                If ShelfDim3 > ShelfDim2 Then
                    strdim3switch = ShelfDim3
                    strdim2switch = ShelfDim2
                    ShelfDim2 = strdim3switch
                    ShelfDim3 = strdim2switch
                End If

                If ShelfDim2 > ShelfDim1 Then
                    strdim2switch = ShelfDim2
                    strdim1switch = ShelfDim1
                    ShelfDim1 = strdim2switch
                    ShelfDim2 = strdim1switch
                End If
	
                If ShelfDim3 > ShelfDim2 Then
                    strdim3switch = ShelfDim3
                    strdim2switch = ShelfDim2
                    ShelfDim2 = strdim3switch
                    ShelfDim3 = strdim2switch
                End If
          
                ReDim Preserve MyArray3(8, y)

                MyArray3(0, y) = myReader("LocationSizeCode")
                MyArray3(1, y) = myReader("OpeningWidth")
                MyArray3(2, y) = myReader("OpeningHeight")
                MyArray3(3, y) = myReader("OpeningDepth")
                MyArray3(4, y) = myReader("WhseArea")
                MyArray3(5, y) = myReader("OpeningWidth") * myReader("OpeningHeight") * myReader("OpeningDepth")
                MyArray3(6, y) = ShelfDim1
                MyArray3(7, y) = ShelfDim2
                MyArray3(8, y) = ShelfDim3
                y = y + 1
                SuperBulkShelfCounter = SuperBulkShelfCounter + 1
                
            Loop
            
        End If
        mySqlConnection.Close()
        
      b = 0
        
        'Sort the array
        For p = y - 1 To 0 Step -1
            For j = y - 2 To 0 Step -1
                If MyArray3(5, j) < MyArray3(5, j + 1) Then
                    For m = 0 To 8
                        s = MyArray3(m, j + 1)
                        MyArray3(m, j + 1) = MyArray3(m, j)
                        MyArray3(m, j) = s
                    Next
                End If
            Next
        Next
        
        FirstSuperBulkShelf = MyArray3(0, 0)
        
        For i = y - 1 To 0 Step -1
            csvheaderstring = csvheaderstring & "," & (MyArray3(0, i))
        Next
        
        'Data from xls file - join with Super Bulk Bases and Super Duper Bulk Part Numbers
        mySqlCommand = New SqlCommand("SELECT PartData.*,SuperBulkBases.*,SuperDuperBulk.* FROM (PartData Left join SuperBulkBases on PartData.PartBase=SuperBulkBases.SuperBulkBase) Left Join SuperDuperBulk on PartData.PartNumber=SuperDuperBulk.[Part Number]", mySqlConnection)
        strCSV3 = "Part Number,Prefix,Base,Suffix,Material Number,Material Type,Width,Height,Depth,Weight,SIL,Inventory On Hand,Location Capacity,Qty > Location Capacity,Pallet Qty,Overstock Pallets Needed,Location,Area" & csvheaderstring & " " & vbCrLf

   
        mySqlConnection.Open()
        myReader = mySqlCommand.ExecuteReader()
        If myReader.HasRows = True Then
            
            Do While myReader.Read()
           
                PartNumber = Class1.trapnulls(myReader("PartNumber"))
                Avgdemand = Class1.trapnulls(myReader("Avgdemand"))
                PlanoSite = Class1.trapnulls(myReader("Site"))
                MaterialNumber = Class1.trapnulls(myReader("MaterialNumber"))
                PackageWidth = Class1.trapnulls(myReader("Width"))
                PackageHeight = Class1.trapnulls(myReader("Height"))
                PackageDepth = Class1.trapnulls(myReader("Depth"))
                If PackageWidth > 0 And PackageHeight > 0 And PackageDepth < 1 Then
                    PackageDepth = 12.7
                End If
                MaterialType = Class1.trapnulls(myReader("MaterialType"))
                MaterialBrand = Class1.trapnulls(myReader("MaterialBrand"))
                PartClass = Class1.trapnulls(myReader("PartClass"))
                Weight = Class1.trapnulls(myReader("Weight"))
                FreeStock = Class1.trapnulls(myReader("FreeStock"))
                SIL = Class1.trapnulls(myReader("SIL"))
                OriginalSIL = Class1.trapnulls(myReader("SIL"))
                PartPrefix = Class1.trapnulls(myReader("PartPrefix"))
                PartBase = Class1.trapnulls(myReader("PartBase"))
                PartSuffix = Class1.trapnulls(myReader("PartSuffix"))
                Zone = Class1.trapnulls(myReader("Zone"))
                Location = Class1.trapnulls(myReader("Location"))
                InventoryOnHand = Class1.trapnulls(myReader("InventoryOnHand"))
                SuperBulkBase = Class1.trapnulls(myReader("Description"))
                SuperDuperBulk = Class1.trapnulls(myReader("SuperDuperBulk"))
                
               
                 
             MultSale = Class1.trapnulls(myReader("MultSale"))
                If MultSale > 0 Then
                    SIL = SIL / MultSale
                End If
                
                Hazmat = Class1.trapnulls(myReader("Hazmat"))
                
                PkgDim1 = PackageWidth
                PkgDim2 = PackageHeight
                PkgDim3 = PackageDepth
                
                If PkgDim3 > PkgDim2 Then
                    strdim3switch = PkgDim3
                    strdim2switch = PkgDim2
                    PkgDim2 = strdim3switch
                    PkgDim3 = strdim2switch
                End If

                If PkgDim2 > PkgDim1 Then
                    strdim2switch = PkgDim2
                    strdim1switch = PkgDim1
                    PkgDim1 = strdim2switch
                    PkgDim2 = strdim1switch
                End If
	
                If PkgDim3 > PkgDim2 Then
                    strdim3switch = PkgDim3
                    strdim2switch = PkgDim2
                    PkgDim2 = strdim3switch
                    PkgDim3 = strdim2switch
                End If
                
                PackagePerimeter = PackageWidth + PackageHeight + PackageDepth
                PartVolume = PackageWidth * PackageHeight * PackageDepth * 1.15
                PartVolumeSIL = PackageWidth * PackageHeight * PackageDepth * SIL * 0.85
                
                If PackageWidth = 0 Or PackageHeight = 0 Or Weight = 0 Then
                    
                   strCSV3 = strCSV3 & PartNumber & "," & PartPrefix & "," & PartBase & "," & PartSuffix & "," & MaterialNumber & "," & MaterialType & "," & PackageWidth & "," & PackageHeight & "," & PackageDepth & "," & Weight & "," & SIL & ",0,0,0,0,0,Missing Data: Width or Height or Depth or Weight"
                    For i = 1 To MezzanineShelfCounter
                        strCSV3 = strCSV3 & ",0"
                    Next
                   
                    For i = 1 To BulkShelfCounter
                        strCSV3 = strCSV3 & ",0"
                    Next
                        
                    For i = 1 To MouldingShelfCounter
                        strCSV3 = strCSV3 & ",0"
                    Next
                        
                    For i = 1 To SuperBulkShelfCounter
                        strCSV3 = strCSV3 & ",0"
                    Next
                    
                    strCSV3 = strCSV3 & vbCrLf

                Else
                    
                    If PackagePerimeter > 622 Or Weight > 3.4 Then
                       
                        If SuperBulkBase = "SuperBulk" Then
                           
                            'Check for SuperDuperBulk
                            If SuperDuperBulk = "415" Then
                                MyArrayMaxFits(1) = Math.Floor(1500 / PkgDim3) * Math.Floor(1750 / PkgDim2) * Math.Floor(3600 / PkgDim1)
                                MyArrayMaxFits(2) = Math.Floor(1500 / PkgDim3) * Math.Floor(3600 / PkgDim2) * Math.Floor(1750 / PkgDim1)
                                MyArrayMaxFits(3) = Math.Floor(1750 / PkgDim3) * Math.Floor(1500 / PkgDim2) * Math.Floor(3600 / PkgDim1)
                                MyArrayMaxFits(4) = Math.Floor(1750 / PkgDim3) * Math.Floor(3600 / PkgDim2) * Math.Floor(1500 / PkgDim1)
                                MyArrayMaxFits(5) = Math.Floor(3600 / PkgDim3) * Math.Floor(1500 / PkgDim2) * Math.Floor(1750 / PkgDim1)
                                MyArrayMaxFits(6) = Math.Floor(3600 / PkgDim3) * Math.Floor(1750 / PkgDim2) * Math.Floor(1500 / PkgDim1)
                                MaxFit = Class1.MaxValOfIntArray(MyArrayMaxFits)
                                
                                If SIL > MaxFit Or InventoryOnHand > MaxFit Then
                                    
                                
                                    If (InventoryOnHand > SIL) And (InventoryOnHand > MaxFit) Then
                                        
                                        Overstock = InventoryOnHand - MaxFit
                                        OverstockVolume = Overstock * PartVolume
                                        OverstockPalletsNeeded = Overstock / MaxFit
                                    Else
                                        Overstock = SIL - MaxFit
                                        OverstockVolume = Overstock * PartVolume
                                        OverstockPalletsNeeded = Overstock / MaxFit
                                    End If
                                    
                                Else
                                    Overstock = 0
                                    OverstockVolume = 0
                                    OverstockPalletsNeeded = 0
                                End If
                                strCSV3 = strCSV3 & PartNumber & "," & PartPrefix & "," & PartBase & "," & PartSuffix & "," & MaterialNumber & "," & MaterialType & "," & PackageWidth & "," & PackageHeight & "," & PackageDepth & "," & Weight & "," & OriginalSIL & "," & InventoryOnHand & "," & SizingQty & "," & Overstock & "," & SizingQty & "," & OverstockPalletsNeeded & ",415,Super Duper Bulk"
                               
                               
                                For i = 1 To MezzanineShelfCounter
                                    strCSV3 = strCSV3 & ",0"
                                Next
                   
                                For i = 1 To BulkShelfCounter
                                    strCSV3 = strCSV3 & ",0"
                                Next
                        
                                For i = 1 To MouldingShelfCounter
                                    strCSV3 = strCSV3 & ",0"
                                Next
                        
                                For i = 1 To SuperBulkShelfCounter
                                    strCSV3 = strCSV3 & ",0"
                                Next
                                
                                strCSV3 = strCSV3 & MaxFit & "," & vbCrLf
                               
                             
                       
                            Else
                                
                            
                            
                            
                                'Super Bulk Location Routine                       
                                MezzanineLocation = FirstSuperBulkShelf
                                strExtraShelfSpace = 0
                                strPreviousShelf = SIL - 1
                                For t = y - 1 To 0 Step -1
                            
                                    MyArrayMaxFits(1) = Math.Floor(MyArray3(6, t) / PkgDim3) * Math.Floor(MyArray3(7, t) / PkgDim2) * Math.Floor(MyArray3(8, t) / PkgDim1)
                                    MyArrayMaxFits(2) = Math.Floor(MyArray3(6, t) / PkgDim3) * Math.Floor(MyArray3(7, t) / PkgDim1) * Math.Floor(MyArray3(8, t) / PkgDim2)
                                    MyArrayMaxFits(3) = Math.Floor(MyArray3(7, t) / PkgDim3) * Math.Floor(MyArray3(6, t) / PkgDim2) * Math.Floor(MyArray3(8, t) / PkgDim1)
                                    MyArrayMaxFits(4) = Math.Floor(MyArray3(7, t) / PkgDim3) * Math.Floor(MyArray3(8, t) / PkgDim2) * Math.Floor(MyArray3(6, t) / PkgDim1)
                                    MyArrayMaxFits(5) = Math.Floor(MyArray3(8, t) / PkgDim3) * Math.Floor(MyArray3(6, t) / PkgDim2) * Math.Floor(MyArray3(7, t) / PkgDim1)
                                    MyArrayMaxFits(6) = Math.Floor(MyArray3(8, t) / PkgDim3) * Math.Floor(MyArray3(7, t) / PkgDim2) * Math.Floor(MyArray3(6, t) / PkgDim1)
                                    MaxFit = Class1.MaxValOfIntArray(MyArrayMaxFits)
                            
                                    If SIL > strPreviousShelf And SIL <= MaxFit Then
                                        MezzanineLocation = MyArray3(0, t)
                                        SizingQty = MaxFit
                                        If MezzanineLocation = "412" Then
                                            SuperBulkPalletVolume = 2125920000
                                        ElseIf MezzanineLocation = "413" Then
                                            SuperBulkPalletVolume = 2768640000
                                        Else
                                            SuperBulkPalletVolume = 3584400000
                                        End If
                                    
                                        If (InventoryOnHand > SIL) And (InventoryOnHand > MaxFit) Then
                                        
                                            Overstock = InventoryOnHand - MaxFit
                                            OverstockVolume = Overstock * PartVolume
                                            OverstockPalletsNeeded = Overstock / MaxFit
                                        Else
                                            Overstock = 0
                                            OverstockVolume = 0
                                            OverstockPalletsNeeded = 0
                                        End If
                                    End If
                            
                                    MyArrayDisplaySizing(w) = MaxFit
                                    w = w + 1
                                    strPreviousShelf = MaxFit
                                    LargestLocation = MyArray3(0, t)
                                
                                Next
                            
                                If SIL > MaxFit Then
                                    MezzanineLocation = LargestLocation
                                    SizingQty = MaxFit
                                    'Overstock calculations
                            
                                    'check to see if Inventory On Hand is Greater than SIL
                                    If InventoryOnHand > MaxFit Or SIL > MaxFit Then
                                        If InventoryOnHand > SIL Then
                                            Overstock = InventoryOnHand - MaxFit
                                        Else
                                            Overstock = SIL - MaxFit
                                        End If
                                    Else
                                        Overstock = 0
                                        OverstockVolume = 0
                                        OverstockPalletsNeeded = 0
                                    End If
                                    
                            
                                   OverstockPalletsNeeded = Overstock / MaxFit
                                Else
                             
                                End If
                                
                                strCSV3 = strCSV3 & PartNumber & "," & PartPrefix & "," & PartBase & "," & PartSuffix & "," & MaterialNumber & "," & MaterialType & "," & PackageWidth & "," & PackageHeight & "," & PackageDepth & "," & Weight & "," & OriginalSIL & "," & InventoryOnHand & "," & SizingQty & "," & Overstock & "," & SizingQty & "," & OverstockPalletsNeeded & "," & MezzanineLocation & ",Super Bulk Area"
                                
                                For i = 1 To MezzanineShelfCounter
                                    strCSV3 = strCSV3 & ",0"
                                Next
                   
                                For i = 1 To BulkShelfCounter
                                    strCSV3 = strCSV3 & ",0"
                                Next
                        
                                For i = 1 To MouldingShelfCounter
                                    strCSV3 = strCSV3 & ",0"
                                Next
                        
                               
                               

                            
                                For w = 0 To SuperBulkShelfCounter - 1
                                    strCSV3 = strCSV3 & "," & MyArrayDisplaySizing(w)
                                Next
                        
                                'strCSV3 = strCSV3 & "," & "0"
                                strCSV3 = strCSV3 & vbCrLf
                       
                                w = 0
                                
                            End If
                            
                            
                        ElseIf SuperBulkBase = "Moulding" Then
                            
                            'Super Bulk Location Routine                       
                            MezzanineLocation = FirstMouldingShelf
                            strExtraShelfSpace = 0
                            strPreviousShelf = SIL - 1
                            For t = z - 1 To 0 Step -1
                            
                                MyArrayMaxFits(1) = Math.Floor(MyArray4(6, t) / PkgDim3) * Math.Floor(MyArray4(7, t) / PkgDim2) * Math.Floor(MyArray4(8, t) / PkgDim1)
                                MyArrayMaxFits(2) = Math.Floor(MyArray4(6, t) / PkgDim3) * Math.Floor(MyArray4(7, t) / PkgDim1) * Math.Floor(MyArray4(8, t) / PkgDim2)
                                MyArrayMaxFits(3) = Math.Floor(MyArray4(7, t) / PkgDim3) * Math.Floor(MyArray4(6, t) / PkgDim2) * Math.Floor(MyArray4(8, t) / PkgDim1)
                                MyArrayMaxFits(4) = Math.Floor(MyArray4(7, t) / PkgDim3) * Math.Floor(MyArray4(8, t) / PkgDim2) * Math.Floor(MyArray4(6, t) / PkgDim1)
                                MyArrayMaxFits(5) = Math.Floor(MyArray4(8, t) / PkgDim3) * Math.Floor(MyArray4(6, t) / PkgDim2) * Math.Floor(MyArray4(7, t) / PkgDim1)
                                MyArrayMaxFits(6) = Math.Floor(MyArray4(8, t) / PkgDim3) * Math.Floor(MyArray4(7, t) / PkgDim2) * Math.Floor(MyArray4(6, t) / PkgDim1)
                                MaxFit = Class1.MaxValOfIntArray(MyArrayMaxFits)
                            
                                If SIL > strPreviousShelf And SIL <= MaxFit Then
                                    MezzanineLocation = MyArray4(0, t)
                                    SizingQty = MaxFit
                                    If (InventoryOnHand > SIL) And (InventoryOnHand > MaxFit) Then
                                        
                                       
                                        
                                        
                                        Overstock = InventoryOnHand - MaxFit
                                        OverstockVolume = Overstock * PartVolume
                                        OverstockPalletsNeeded = Overstock / MaxFit
                                    Else
                                        Overstock = 0
                                        OverstockVolume = 0
                                        OverstockPalletsNeeded = 0
                                    End If
                                End If
                            
                                MyArrayDisplaySizing(w) = MaxFit
                                w = w + 1
                                strPreviousShelf = MaxFit
                                
                                LargestLocation = MyArray4(0, t)
                                
                            Next
                            
                            If SIL > MaxFit Then
                                MezzanineLocation = LargestLocation
                                SizingQty = MaxFit
                                'Overstock calculations
                                If InventoryOnHand > MaxFit Or SIL > MaxFit Then
                                    If InventoryOnHand > SIL Then
                                        Overstock = InventoryOnHand - MaxFit
                                    Else
                                        Overstock = SIL - MaxFit
                                    End If
                                
                                    OverstockVolume = Overstock * PartVolume
                                    OverstockPalletsNeeded = Overstock / MaxFit
                                Else
                                    Overstock = 0
                                    OverstockVolume = 0
                                    OverstockPalletsNeeded = 0
                                End If
                                
                            Else
                            End If
                            
                            strCSV3 = strCSV3 & PartNumber & "," & PartPrefix & "," & PartBase & "," & PartSuffix & "," & MaterialNumber & "," & MaterialType & "," & PackageWidth & "," & PackageHeight & "," & PackageDepth & "," & Weight & "," & OriginalSIL & "," & InventoryOnHand & "," & SizingQty & "," & Overstock & "," & SizingQty & "," & OverstockPalletsNeeded & "," & MezzanineLocation & ",Moulding"
                            
                            
                            For i = 1 To MezzanineShelfCounter
                                strCSV3 = strCSV3 & ",0"
                            Next
                   
                            For i = 1 To BulkShelfCounter
                                strCSV3 = strCSV3 & ",0"
                            Next
                        
                           

                            
                            For w = 0 To MouldingShelfCounter - 1
                                strCSV3 = strCSV3 & "," & MyArrayDisplaySizing(w)
                            Next
                            
                            For i = 1 To SuperBulkShelfCounter
                                strCSV3 = strCSV3 & ",0"
                            Next
                            
                            strCSV3 = strCSV3 & vbCrLf
                            w = 0
                            
                        Else
                            
                           
                            ' Bulk Location Routine                       
                            MezzanineLocation = FirstBulkShelf
                            strExtraShelfSpace = 0
                            strPreviousShelf = SIL - 1
                            For t = x - 1 To 0 Step -1
                            
                                MyArrayMaxFits(1) = Math.Floor(MyArray2(6, t) / PkgDim3) * Math.Floor(MyArray2(7, t) / PkgDim2) * Math.Floor(MyArray2(8, t) / PkgDim1)
                                MyArrayMaxFits(2) = Math.Floor(MyArray2(6, t) / PkgDim3) * Math.Floor(MyArray2(7, t) / PkgDim1) * Math.Floor(MyArray2(8, t) / PkgDim2)
                                MyArrayMaxFits(3) = Math.Floor(MyArray2(7, t) / PkgDim3) * Math.Floor(MyArray2(6, t) / PkgDim2) * Math.Floor(MyArray2(8, t) / PkgDim1)
                                MyArrayMaxFits(4) = Math.Floor(MyArray2(7, t) / PkgDim3) * Math.Floor(MyArray2(8, t) / PkgDim2) * Math.Floor(MyArray2(6, t) / PkgDim1)
                                MyArrayMaxFits(5) = Math.Floor(MyArray2(8, t) / PkgDim3) * Math.Floor(MyArray2(6, t) / PkgDim2) * Math.Floor(MyArray2(7, t) / PkgDim1)
                                MyArrayMaxFits(6) = Math.Floor(MyArray2(8, t) / PkgDim3) * Math.Floor(MyArray2(7, t) / PkgDim2) * Math.Floor(MyArray2(6, t) / PkgDim1)
                                MaxFit = Class1.MaxValOfIntArray(MyArrayMaxFits)
                            
                                If SIL > strPreviousShelf And SIL <= MaxFit Then
                                    MezzanineLocation = MyArray2(0, t)
                                    SizingQty = MaxFit
                                    
                                    MyArrayMaxFitsMezzanine(1) = Math.Floor(1200 / PkgDim3) * Math.Floor(1220 / PkgDim2) * Math.Floor(1300 / PkgDim1)
                                    MyArrayMaxFitsMezzanine(2) = Math.Floor(1200 / PkgDim3) * Math.Floor(1220 / PkgDim1) * Math.Floor(1300 / PkgDim2)
                                    MyArrayMaxFitsMezzanine(3) = Math.Floor(1220 / PkgDim3) * Math.Floor(1200 / PkgDim2) * Math.Floor(1300 / PkgDim1)
                                    MyArrayMaxFitsMezzanine(4) = Math.Floor(1220 / PkgDim3) * Math.Floor(1300 / PkgDim2) * Math.Floor(1200 / PkgDim1)
                                    MyArrayMaxFitsMezzanine(5) = Math.Floor(1300 / PkgDim3) * Math.Floor(1200 / PkgDim2) * Math.Floor(1220 / PkgDim1)
                                    MyArrayMaxFitsMezzanine(6) = Math.Floor(1300 / PkgDim3) * Math.Floor(1220 / PkgDim2) * Math.Floor(1200 / PkgDim1)
                                    PalletMaxFit = Class1.MaxValOfIntArray(MyArrayMaxFitsMezzanine)
                                    
                                    
                                    If (InventoryOnHand > SIL) And (InventoryOnHand > MaxFit) Then
                                        Overstock = InventoryOnHand - MaxFit
                                        OverstockVolume = Overstock * PartVolume
                                        OverstockPalletsNeeded = Overstock / PalletMaxFit
                                    Else
                                        Overstock = 0
                                        OverstockVolume = 0
                                        OverstockPalletsNeeded = 0
                                    End If
                                End If
                            
                                'strCSV3 = strCSV3 & MaxFit & ","
                                MyArrayDisplaySizing(w) = MaxFit
                                w = w + 1
                                strPreviousShelf = MaxFit
                                LargestLocation = MyArray2(0, t)
                                
                            Next
                            
                           
                            
                            If SIL > MaxFit Then
                                MezzanineLocation = LargestLocation
                                SizingQty = MaxFit
                                
                                MyArrayMaxFitsMezzanine(1) = Math.Floor(940 / PkgDim3) * Math.Floor(1067 / PkgDim2) * Math.Floor(1219 / PkgDim1)
                                MyArrayMaxFitsMezzanine(2) = Math.Floor(940 / PkgDim3) * Math.Floor(1219 / PkgDim2) * Math.Floor(1067 / PkgDim1)
                                MyArrayMaxFitsMezzanine(3) = Math.Floor(1067 / PkgDim3) * Math.Floor(940 / PkgDim2) * Math.Floor(1219 / PkgDim1)
                                MyArrayMaxFitsMezzanine(4) = Math.Floor(1067 / PkgDim3) * Math.Floor(1219 / PkgDim2) * Math.Floor(940 / PkgDim1)
                                MyArrayMaxFitsMezzanine(5) = Math.Floor(1219 / PkgDim3) * Math.Floor(940 / PkgDim2) * Math.Floor(1067 / PkgDim1)
                                MyArrayMaxFitsMezzanine(6) = Math.Floor(1219 / PkgDim3) * Math.Floor(1067 / PkgDim2) * Math.Floor(940 / PkgDim1)
                                PalletMaxFit = Class1.MaxValOfIntArray(MyArrayMaxFitsMezzanine)
                                'Overstock calculations
                                If InventoryOnHand > MaxFit Or SIL > MaxFit Then
                                    If InventoryOnHand > SIL Then
                                        Overstock = InventoryOnHand - MaxFit
                                    Else
                                        Overstock = SIL - MaxFit
                                    End If
                                
                                    OverstockVolume = Overstock * PartVolume
                                    OverstockPalletsNeeded = Overstock / PalletMaxFit
                                Else
                                    Overstock = 0
                                    OverstockVolume = 0
                                    OverstockPalletsNeeded = 0
                                End If
                                
                            Else
                             End If
                            
                            strCSV3 = strCSV3 & PartNumber & "," & PartPrefix & "," & PartBase & "," & PartSuffix & "," & MaterialNumber & "," & MaterialType & "," & PackageWidth & "," & PackageHeight & "," & PackageDepth & "," & Weight & "," & OriginalSIL & "," & InventoryOnHand & "," & SizingQty & "," & Overstock & "," & PalletMaxFit & "," & OverstockPalletsNeeded & "," & MezzanineLocation & ",Bulk Area"
                            
                            For i = 1 To MezzanineShelfCounter
                                strCSV3 = strCSV3 & ",0"
                            Next
                           

                            For w = 0 To BulkShelfCounter - 1
                                
                                strCSV3 = strCSV3 & "," & MyArrayDisplaySizing(w)
                            Next
                            
                            For i = 1 To MouldingShelfCounter
                                strCSV3 = strCSV3 & ",0"
                            Next
                            
                            For i = 1 To SuperBulkShelfCounter
                                strCSV3 = strCSV3 & ",0"
                            Next
                        
                            strCSV3 = strCSV3 & vbCrLf
                        
                            w = 0
                        End If
                        
                        
                    Else

                    
                      

                        MezzanineLocation = FirstMezzanineShelf
                        strExtraShelfSpace = 0
                        strPreviousShelf = SIL - 1
                        
                        Overstock = 0
                        OverstockVolume = 0
                        OverstockPalletsNeeded = 0
                        
                        
                        For t = k - 1 To 0 Step -1
                            
                       
                            MyArrayMaxFits(1) = Math.Floor(MyArray(6, t) / PkgDim3) * Math.Floor(MyArray(7, t) / PkgDim2) * Math.Floor(MyArray(8, t) / PkgDim1)
                            MyArrayMaxFits(2) = Math.Floor(MyArray(6, t) / PkgDim3) * Math.Floor(MyArray(8, t) / PkgDim2) * Math.Floor(MyArray(7, t) / PkgDim1)
                            MyArrayMaxFits(3) = Math.Floor(MyArray(7, t) / PkgDim3) * Math.Floor(MyArray(6, t) / PkgDim2) * Math.Floor(MyArray(8, t) / PkgDim1)
                            MyArrayMaxFits(4) = Math.Floor(MyArray(7, t) / PkgDim3) * Math.Floor(MyArray(8, t) / PkgDim2) * Math.Floor(MyArray(6, t) / PkgDim1)
                            MyArrayMaxFits(5) = Math.Floor(MyArray(8, t) / PkgDim3) * Math.Floor(MyArray(6, t) / PkgDim2) * Math.Floor(MyArray(7, t) / PkgDim1)
                            MyArrayMaxFits(6) = Math.Floor(MyArray(8, t) / PkgDim3) * Math.Floor(MyArray(7, t) / PkgDim2) * Math.Floor(MyArray(6, t) / PkgDim1)
                            MaxFit = Class1.MaxValOfIntArray(MyArrayMaxFits)
                            
                            If MaxFit = MyArrayMaxFits(1) Then
                                MaxFitsOrientation = "1"
                            ElseIf MaxFit = MyArrayMaxFits(2) Then
                                MaxFitsOrientation = "2"
                            ElseIf MaxFit = MyArrayMaxFits(3) Then
                                MaxFitsOrientation = "3"
                            ElseIf MaxFit = MyArrayMaxFits(4) Then
                                MaxFitsOrientation = "4"
                            ElseIf MaxFit = MyArrayMaxFits(5) Then
                                MaxFitsOrientation = "5"
                            Else
                                MaxFitsOrientation = "6"
                            
                            End If
                            
                            If SIL > strPreviousShelf And SIL <= MaxFit Then
                                MezzanineLocation = MyArray(0, t)
                                SizingQty = MaxFit
                                
                                MyArrayMaxFitsMezzanine(1) = Math.Floor(940 / PkgDim3) * Math.Floor(1067 / PkgDim2) * Math.Floor(1219 / PkgDim1)
                                MyArrayMaxFitsMezzanine(2) = Math.Floor(940 / PkgDim3) * Math.Floor(1067 / PkgDim1) * Math.Floor(1219 / PkgDim2)
                                MyArrayMaxFitsMezzanine(3) = Math.Floor(1067 / PkgDim3) * Math.Floor(940 / PkgDim2) * Math.Floor(1219 / PkgDim1)
                                MyArrayMaxFitsMezzanine(4) = Math.Floor(1067 / PkgDim3) * Math.Floor(1219 / PkgDim2) * Math.Floor(940 / PkgDim1)
                                MyArrayMaxFitsMezzanine(5) = Math.Floor(1219 / PkgDim3) * Math.Floor(940 / PkgDim2) * Math.Floor(1067 / PkgDim1)
                                MyArrayMaxFitsMezzanine(6) = Math.Floor(1219 / PkgDim3) * Math.Floor(1067 / PkgDim2) * Math.Floor(940 / PkgDim1)
                                PalletMaxFit = Class1.MaxValOfIntArray(MyArrayMaxFitsMezzanine)
                                
                                
                                If (InventoryOnHand > SIL) And (InventoryOnHand > MaxFit) Then
                                    Overstock = InventoryOnHand - MaxFit
                                    OverstockVolume = Overstock * PartVolume
                                    OverstockPalletsNeeded = Overstock / PalletMaxFit
                                Else
                                    Overstock = 0
                                    OverstockVolume = 0
                                    OverstockPalletsNeeded = 0
                                End If
                                    
                            End If
                           
                            MyArrayDisplaySizing(w) = MaxFit
                            w = w + 1
                            
                            
                            strPreviousShelf = MaxFit
                            
                            LargestLocation = MyArray(0, t)
                                
                        Next
                        
                       
                            
                        If SIL > MaxFit Then
                            MezzanineLocation = LargestLocation
                            SizingQty = MaxFit
                            'Overstock calculations
                            
                            MyArrayMaxFitsMezzanine(1) = Math.Floor(940 / PkgDim3) * Math.Floor(1067 / PkgDim2) * Math.Floor(1219 / PkgDim1)
                            MyArrayMaxFitsMezzanine(2) = Math.Floor(940 / PkgDim3) * Math.Floor(1067 / PkgDim1) * Math.Floor(1219 / PkgDim2)
                            MyArrayMaxFitsMezzanine(3) = Math.Floor(1067 / PkgDim3) * Math.Floor(940 / PkgDim2) * Math.Floor(1219 / PkgDim1)
                            MyArrayMaxFitsMezzanine(4) = Math.Floor(1067 / PkgDim3) * Math.Floor(1219 / PkgDim2) * Math.Floor(940 / PkgDim1)
                            MyArrayMaxFitsMezzanine(5) = Math.Floor(1219 / PkgDim3) * Math.Floor(940 / PkgDim2) * Math.Floor(1067 / PkgDim1)
                            MyArrayMaxFitsMezzanine(6) = Math.Floor(1219 / PkgDim3) * Math.Floor(1067 / PkgDim2) * Math.Floor(940 / PkgDim1)
                            PalletMaxFit = Class1.MaxValOfIntArray(MyArrayMaxFitsMezzanine)
                            
                            
                            'check to see if Inventory On Hand is Greater than SIL
                            If InventoryOnHand > MaxFit Or SIL > MaxFit Then
                                If InventoryOnHand > SIL Then
                                    Overstock = InventoryOnHand - MaxFit
                                Else
                                    Overstock = SIL - MaxFit
                                End If
                            
                                OverstockPalletsNeeded = Overstock / PalletMaxFit
                            Else
                                Overstock = 0
                                OverstockVolume = 0
                                OverstockPalletsNeeded = 0
                            End If
                            
                        Else
                        End If
                        
                        strCSV3 = strCSV3 & PartNumber & "," & PartPrefix & "," & PartBase & "," & PartSuffix & "," & MaterialNumber & "," & MaterialType & "," & PackageWidth & "," & PackageHeight & "," & PackageDepth & "," & Weight & "," & OriginalSIL & "," & InventoryOnHand & "," & SizingQty & "," & Overstock & "," & PalletMaxFit & "," & OverstockPalletsNeeded & "," & MezzanineLocation & ",Mezzanine"

                        For w = 0 To MezzanineShelfCounter - 1
                            strCSV3 = strCSV3 & "," & MyArrayDisplaySizing(w)
                        Next
                       For i = 1 To BulkShelfCounter
                            strCSV3 = strCSV3 & ",0"
                        Next
                        
                        For i = 1 To MouldingShelfCounter
                            strCSV3 = strCSV3 & ",0"
                        Next
                        
                        For i = 1 To SuperBulkShelfCounter
                            strCSV3 = strCSV3 & ",0"
                        Next

                        strCSV3 = strCSV3 & vbCrLf
                  
                      MezzanineLocation = ""
                        MezzanineSize = ""
                        w = 0
                        
                    End If
                End If
                
             
                Response.Write(PartNumber & "<br>")
                
            Loop
            
        End If
        mySqlConnection.Close()
       
        Response.Clear()
        Response.AddHeader("content-disposition", "attachment;filename=DataExport.csv")
        Response.Charset = ""
        Response.ContentType = "application/text"
        Response.Write(strCSV3.ToString)
        Response.End()
    End Sub
    </script>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
        <center>
            <asp:Image ID="Image1" runat="server" ImageUrl="images/PageTitleCreatePlanograph.jpg" />
        <div>
        <table cellspacing="20">
        <tr><td colspan="2">To create the Planograph File: <br /> 1. Select a Location Name from the "Select Mezzanine Shelf Sizing" drop down list.<br />2. Click on the "Create CSV" button.</td></tr>
         <tr>
                                                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString1 %>" SelectCommand="Select Distinct LocationName from LocationSizes Where WhseCategory='Mezzanine' order by LocationName" ></asp:SqlDataSource>

						<td align="right">Select Mezzanine Shelf Sizing: </td>
						<td align="left">
						<asp:DropDownList ID="Mezzanine"
                        runat="server" AppendDataBoundItems="true" DataSourceID="SqlDataSource1" DataTextField="LocationName" DataValueField="LocationName" Enabled="true">
                   
                        </asp:DropDownList></td>
                        </tr>
                        
                  
                        <tr><td colspan="2" align="center">
        <asp:ImageButton ID="ImageButton2" runat="server" OnClick="CreateCSV"  ImageUrl="images/CreateCSVButton.png" CausesValidation="false" />  
</td></tr>
</table>
<br /><br /><br /><br /><br /><br /><br /><br /><br />
        </div>
        </center>
</asp:Content>