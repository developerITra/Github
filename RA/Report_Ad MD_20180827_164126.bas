VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Ad MD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function AdditionalTerms() As String
AdditionalTerms = "There will be no abatement of interest in the event additional funds are tendered before settlement or if settlment is delayed for any reason.  " & _
    "The noteholder shall not be obligated to pay interest if it is the purchaser. TIME IS OF THE ESSENCE FOR THE PURCHASER.  " & _
    "All public charges or assessments, to the extent such amounts survive foreclosure, including water/sewer charges, ground rent, agricultural tax, condo/HOA dues, whether incurred " & _
    "prior to or after the sale, and all other costs incident to settlement to be paid by the purchaser.  " & _
    "In the event taxes, any other public charges or condo/HOA fees have been advanced, a credit will be due to the seller, to be adjusted from the date of sale at the time of settlement.  " & _
    "Cost of all documentary stamps, transfer taxes and settlement expenses shall be borne " & _
    "by the purchaser.  Purchaser shall be responsible for obtaining physical possession of the property.  " & _
    "Purchaser assumes the risk of loss or damage to the property from the date of sale forward.  " & _
    "Additional terms to be announced at the time of sale."
End Function

