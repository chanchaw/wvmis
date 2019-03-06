Attribute VB_Name = "modBillObject"
Option Explicit

'Public Sub GetFormObject(ByRef o As Object, ByVal FormName As String)
'    Select Case FormName
'        Case "frmModBL"
'            Set o = New frmModBL
'    End Select
'End Sub


Public Function GetFormNew(ByVal FormName As String) As Object
    Select Case FormName
        Case "frmNavigatorOrder"
            Set GetFormNew = New frmNavigatorOrder
            
        Case "frmNavigatorDT"
            Set GetFormNew = New frmNavigatorDT
            
        Case "frmNavigatorGoid"
            Set GetFormNew = New frmNavigatorGoid
            
            
        Case "frmModBL"
            Set GetFormNew = New frmModBL
            
        Case "frmModBLR"
            Set GetFormNew = New frmModBLR
            
        Case "frmModBLRvsf"
            Set GetFormNew = New frmModBLRvsf
            
        Case "frmModBLS"
            Set GetFormNew = New frmModBLS
            
        Case "frmModBLSEdit"
            Set GetFormNew = New frmModBLSEdit
            
        Case "frmModBLSTree"
            Set GetFormNew = New frmModBLSTree
            
        Case "frmModGoods"
            Set GetFormNew = New frmModGoods
            
        Case "frmModBLROrder"
            Set GetFormNew = New frmModBLROrder
            
        Case "frmNavigatorAccessory"
            Set GetFormNew = New frmNavigatorAccessory
            
        Case "frmStorageQueryAcc"
            Set GetFormNew = New frmStorageQueryAcc
            
        Case "frmStorageQueryYarn"
            Set GetFormNew = New frmStorageQueryYarn
            
        Case "frmStorageQueryWhite"
            Set GetFormNew = New frmStorageQueryWhite
            
        Case "frmStorageQueryColor"
            Set GetFormNew = New frmStorageQueryColor
            
        Case "frmNavigatorYarn"
            Set GetFormNew = New frmNavigatorYarn
            
        Case "frmNavigatorWhite"
            Set GetFormNew = New frmNavigatorWhite
'
        Case "frmNavigatorColor"
            Set GetFormNew = New frmNavigatorColor
            
        Case "frmModBLRAcce"
            Set GetFormNew = New frmModBLRAcce
            
        Case "frmModBLROrderEx"
            Set GetFormNew = New frmModBLROrderEx
            
        Case "frmModBLAccessory"
            Set GetFormNew = New frmModBLAccessory
        
        Case "frmModBLWhite"
            Set GetFormNew = New frmModBLWhite
            
        Case "frmModBLColorCloth"
            Set GetFormNew = New frmModBLColorCloth
            
        Case "frmModBLRAcceSpec"
            Set GetFormNew = New frmModBLRAcceSpec
            
        Case "frmModBLOrder"
            Set GetFormNew = New frmModBLOrder
            
        Case "frmModBLSColor"
            Set GetFormNew = New frmModBLSColor
              
        Case "frmOrder"
            Set GetFormNew = New frmOrder
            
        Case "frmModBLSource"
            Set GetFormNew = New frmModBLSource
            
        Case "frmModBLSColor_Edit"
            Set GetFormNew = New frmModBLSColor_Edit
            
        Case "frmDingDanReport"
            Set GetFormNew = New frmDingDanReport
            
        Case "frmDingDanReport_Edit"
            Set GetFormNew = New frmDingDanReport_Edit
             
        Case "frmOriginalOrder"
            Set GetFormNew = New frmOriginalOrder
            
        Case "frmWhiteOrder"
            Set GetFormNew = New frmWhiteOrder
            
        Case "frmOriginalInset"
            Set GetFormNew = New frmOriginalInset
               
        Case "frmWhiteInset"
            Set GetFormNew = New frmWhiteInset
            
        Case "frmSelectFine"
            Set GetFormNew = New frmSelectFine
            
        Case "frmOriginalcontract"
            Set GetFormNew = New frmOriginalcontract
            
        Case "frmWhitecontract"
            Set GetFormNew = New frmWhitecontract
                     
        Case "frmWhiteProInset"
            Set GetFormNew = New frmWhiteProInset
            
        Case "frmWhitePrint"
            Set GetFormNew = New frmWhitePrint
                   
        Case "frmDingDanSelect"
            Set GetFormNew = New frmDingDanSelect
            
        Case "frmColorPrint"
            Set GetFormNew = New frmColorPrint
         
        Case "frmWhiteSelect"
            Set GetFormNew = New frmWhiteSelect
            
        Case "frmStorageRKCKNewFP"
            Set GetFormNew = New frmStorageRKCKNewFP
         
        Case "frmSelectWhite"
            Set GetFormNew = New frmSelectWhite
         
        Case "frmOriginalSelect"
            Set GetFormNew = New frmOriginalSelect
         
        Case "frmColorOrder"
            Set GetFormNew = New frmColorOrder
            
        Case "frmOriginalOrderInset"
            Set GetFormNew = New frmOriginalOrderInset
         
        Case "frmOriginalReport"
            Set GetFormNew = New frmOriginalReport
         
        Case "frmColorInsetDetail"
            Set GetFormNew = New frmColorInsetDetail
         
        Case "frmWhiteProcessInset"
            Set GetFormNew = New frmWhiteProcessInset
         
        Case "frmWhiteProcessreport"
            Set GetFormNew = New frmWhiteProcessreport
         
        Case "frmwhiteprocess_Sum"
            Set GetFormNew = New frmwhiteprocess_Sum
            
        Case "frmOriginalOrderreport"
            Set GetFormNew = New frmOriginalOrderreport
            
        Case "frmColorReturn"
            Set GetFormNew = New frmColorReturn
         
         Case "frmWhiteFabricReport"
            Set GetFormNew = New frmWhiteFabricReport
         
        Case "frmWhiteProductionReport"
            Set GetFormNew = New frmWhiteProductionReport
            
        Case "frmWhiteDelivery"
            Set GetFormNew = New frmWhiteDelivery
            
        Case "frmColorprocess"
            Set GetFormNew = New frmColorprocess
            
        Case "frmColorDelivery"
            Set GetFormNew = New frmColorDelivery
         
        Case "frmOriginalDelivery"
            Set GetFormNew = New frmOriginalDelivery
         
        Case "frmColorClientReturn"
            Set GetFormNew = New frmColorClientReturn
         
        Case "frmWhiteReturn"
            Set GetFormNew = New frmWhiteReturn
         
        Case "frmOriginalReturn"
            Set GetFormNew = New frmOriginalReturn
         
        Case "frmWhitetransfers"
            Set GetFormNew = New frmWhitetransfers
                 
        Case "frmOriginalOrderInsetReport"
            Set GetFormNew = New frmOriginalOrderInsetReport
         
        Case "frmOriginalinventory"
            Set GetFormNew = New frmOriginalinventory
            
        Case "frmOriginalProducerReport"
            Set GetFormNew = New frmOriginalProducerReport
            
        Case "frmOriginalSumReport"
            Set GetFormNew = New frmOriginalSumReport
            
        Case "frmOriginalprocessReturn"
            Set GetFormNew = New frmOriginalprocessReturn
            
        Case "frmOriginaltransfers"
            Set GetFormNew = New frmOriginaltransfers
            
        Case "frmOriginalprocessReport"
            Set GetFormNew = New frmOriginalprocessReport
            
        Case "frmWhiteReport"
            Set GetFormNew = New frmWhiteReport
            
        Case "frmWhiteDeliveryReport"
            Set GetFormNew = New frmWhiteDeliveryReport
            
        Case "frmWhiteSumRepoart"
            Set GetFormNew = New frmWhiteSumRepoart
            
        Case "frmWhiteinventoryReport"
            Set GetFormNew = New frmWhiteinventoryReport
            
        Case "frmColorInsetReport"
            Set GetFormNew = New frmColorInsetReport
            
        Case "frmColorDeliveryReport"
            Set GetFormNew = New frmColorDeliveryReport
            
        Case "frmColorRepair"
            Set GetFormNew = New frmColorRepair
                      
        Case "frmColorReturnRepair"
            Set GetFormNew = New frmColorReturnRepair
            
        Case "frmColorRepairReport"
            Set GetFormNew = New frmColorRepairReport
            
        Case "frmOriginalProducerDetailReport"
            Set GetFormNew = New frmOriginalProducerDetailReport
            
        Case "frmOriginalLastYear"
            Set GetFormNew = New frmOriginalLastYear
            
        Case "frmWhitelastyear"
            Set GetFormNew = New frmWhitelastyear
            
        Case "frmColorlastyear"
            Set GetFormNew = New frmColorlastyear
            
        Case "frmWhiteProcureInset"
            Set GetFormNew = New frmWhiteProcureInset
            
        Case "frmWhiteProcureReturn"
            Set GetFormNew = New frmWhiteProcureReturn
            
        Case "frmWhiteProcureReport"
            Set GetFormNew = New frmWhiteProcureReport
            
        Case "frmComposition"
            Set GetFormNew = New frmComposition
            
        Case "frmOriginalLastYearReport"
            Set GetFormNew = New frmOriginalLastYearReport
            
        Case "frmWhitelastyearReport"
            Set GetFormNew = New frmWhitelastyearReport
            
        Case "frmColorlastyearReport"
            Set GetFormNew = New frmColorlastyearReport
            
        Case "frmStorageQueryMetals"
            Set GetFormNew = New frmStorageQueryMetals
            
        Case "frmColorinventoryReport"
            Set GetFormNew = New frmColorinventoryReport
            
        Case "frmWhiteCurlyPrint"
            Set GetFormNew = New frmWhiteCurlyPrint
            
        Case "frmColorJRKprint"
            Set GetFormNew = New frmColorJRKprint
            
        Case "frmColorJRKprintDetail"
            Set GetFormNew = New frmColorJRKprintDetail
            
        Case "frmColorOutDelivery"
            Set GetFormNew = New frmColorOutDelivery
            
        Case "frmColorProcessReport"
            Set GetFormNew = New frmColorProcessReport
            
        Case "frmColorprocessReturn"
            Set GetFormNew = New frmColorprocessReturn
            
        Case "frmColorProcessRepair"
            Set GetFormNew = New frmColorProcessRepair
            
        Case "frmColorOutfactory"
            Set GetFormNew = New frmColorOutfactory
            
        Case "frmColorfactory"
            Set GetFormNew = New frmColorfactory
            
        Case "fmrColorSelect"
            Set GetFormNew = New fmrColorSelect
            
        Case "frmColorProcure"
            Set GetFormNew = New frmColorProcure
            
        Case "frmColorProcureDelivery"
            Set GetFormNew = New frmColorProcureDelivery
            
        Case "frmColorProcureReport"
            Set GetFormNew = New frmColorProcureReport
            
        Case "frmColorfactoryReport"
            Set GetFormNew = New frmColorfactoryReport
        
        Case "frmOrderColorReport"
            Set GetFormNew = New frmOrderColorReport
            
        Case "frmOrderWhiteReport"
            Set GetFormNew = New frmOrderWhiteReport
            
        Case "frmSelectXMD"
            Set GetFormNew = New frmSelectXMD
            
        Case "frmOrderProduct"
            Set GetFormNew = New frmOrderProduct
        
        Case "frmOfficeDetails"
            Set GetFormNew = New frmOfficeDetails
            
       Case "frmOfficeCategory"
            Set GetFormNew = New frmOfficeCategory
            
       Case "frmfreight"
            Set GetFormNew = New frmfreight
        
        Case "frmNavigatorfinancial"
            Set GetFormNew = New frmNavigatorfinancial
            
        Case "frmSchedule"
            Set GetFormNew = New frmSchedule
            
        Case "frmProgress"
            Set GetFormNew = New frmProgress
        
        Case "frmCost"
            Set GetFormNew = New frmCost
        
        Case "frmColorProcessReport_2"
            Set GetFormNew = New frmColorProcessReport_2
        
        Case "frmNavigatorCP"
            Set GetFormNew = New frmNavigatorCP
        
        Case "frmProductDeliveryReport"
            Set GetFormNew = New frmProductDeliveryReport

        Case "frmProductProcureReport"
            Set GetFormNew = New frmProductProcureReport
            
        Case "frmOriginalPositionReport"
            Set GetFormNew = New frmOriginalPositionReport
            
        Case "frmOriginalPosition"
            Set GetFormNew = New frmOriginalPosition
            
        Case "frmWhitePosition"
            Set GetFormNew = New frmWhitePosition
            
        Case "frmWhitePositionReport"
            Set GetFormNew = New frmWhitePositionReport
            
        Case "frmColorPosition"
            Set GetFormNew = New frmColorPosition
            
        Case "frmColorPositionReport"
            Set GetFormNew = New frmColorPositionReport
            
        Case "frmProductPosition"
            Set GetFormNew = New frmProductPosition
            
        Case "frmProductPositionReport"
            Set GetFormNew = New frmProductPositionReport
            
        Case "frmColorSumRepoart"
            Set GetFormNew = New frmColorSumRepoart
            
        Case "frmWhiteSumRepoart_White"
            Set GetFormNew = New frmWhiteSumRepoart_White
            
        Case "frmWhiteSumRepoart_ALL"
            Set GetFormNew = New frmWhiteSumRepoart_ALL
            
        Case "frmYarnSumRepoart_ALL"
            Set GetFormNew = New frmYarnSumRepoart_ALL
            
        Case "frmWhiteProduction"
            Set GetFormNew = New frmWhiteProduction
            
      Case "frmColorDJRK"
            Set GetFormNew = New frmColorDJRK
            
    Case "frmOriginalOrderInsetReport_CaiGou"
            Set GetFormNew = New frmOriginalOrderInsetReport_CaiGou
            
            
    End Select
End Function

