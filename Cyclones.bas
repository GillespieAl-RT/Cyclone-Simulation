Attribute VB_Name = "Module1"
Function Plitt_InputsOutputs_2()
'Defines Input and Output table for Plitt Factors
Dim Functions(21, 1)
Functions(0, 0) = "-- Plitt Inputs --"
Functions(0, 1) = "-"
Functions(1, 0) = "d50 correction factor"
Functions(1, 1) = "F"
Functions(2, 0) = "Cyclone diameter (cm)"
Functions(2, 1) = "Dc"
Functions(3, 0) = "Cyclone feed inlet diameter (cm)"
Functions(3, 1) = "Di"
Functions(4, 0) = "Cyclone overflow or Vortex finder diameter (cm)"
Functions(4, 1) = "D_o"
Functions(5, 0) = "Cyclone underflow or Apex diameter (cm)"
Functions(5, 1) = "Du"
Functions(6, 0) = "Free vortex height in cyclone (cm)"
Functions(6, 1) = "h"
Functions(7, 0) = "Solids density at Temperature (t/m3)"
Functions(7, 1) = "rho_s"
Functions(8, 0) = "Liquid density at Temperature (t/m3)"
Functions(8, 1) = "rho_l"
Functions(9, 0) = "Feed slurry density (kg/m3)"
Functions(9, 1) = "rho_f"
Functions(10, 0) = "Pressure Factor"
Functions(10, 1) = "PF"
Functions(11, 0) = "Sharpness Factor"
Functions(11, 1) = "SharpFactor"
Functions(12, 0) = "S-Factor"
Functions(12, 1) = "SFactor"
Functions(13, 0) = "Feed PSD %-45"
Functions(13, 1) = "FeedMinus45"
Functions(14, 0) = "Feed PSD %+150"
Functions(14, 1) = "FeedPlus150"
Functions(15, 0) = "Feed Flow (m3/hr)"
Functions(15, 1) = "Feed_Flow"
Functions(16, 0) = "Feed [solids] (g/L)"
Functions(16, 1) = "Feed_solids"
Functions(17, 0) = "Dilution flow (m3/h)"
Functions(17, 1) = "Dilution_flow"
Functions(18, 0) = "# cyclones"
Functions(18, 1) = "No_cyclones"
Functions(19, 0) = "Feed SPO"
Functions(19, 1) = "FeedSPO"


Plitt_InputsOutputs_2 = Functions
End Function
Function Grade_Efficiency_actual(Grade_Efficiency_Corrected, R_liquid As Double)
'A.Gillespie, 28/8/23
'Actual grade efficiency from corrected grade efficiency
'R_liquid = liquid recovery to underflow
Grade_Efficiency_actual = Grade_Efficiency_Corrected + (R_liquid * (1 - Grade_Efficiency_Corrected))
End Function
Function Grade_Efficiency_Corrected(Grade_Efficiency_actual, R_liquid As Double)
'A.Gillespie, 28/8/23
'Corrected grade efficiency from actual grade efficiency
'R_liquid = liquid recovery to underflow
Grade_Efficiency_Corrected = (Grade_Efficiency_actual - R_liquid) / (1 - R_liquid)
End Function
Function Plitt(F, Dc, di, D_o, Du, h, rho_s, rho_l, rho_f, PF, SharpFactor, SFactor, FeedMinus45, FeedPlus150, Feed_Flow, Feed_Solids, Dilution_Flow, No_cyclones, FeedSPO, Mode As Double)

'A.Gillespie 23/11/23
'Plitt cyclone model
'Refer https://help.syscad.net/Hydrocyclone for model description

'Mode=0 full array; mode=1 underflow; mode=2 overflow; mode=3 pressure drop; mode=4 UF %-53; mode=5 d50

'Datebomb to ensure this version expires
 Dim DateMax As Date
 DateMax = "30/9/2024"
 If Now() < DateMax Then
 Else
 MsgBox "Plitt calculator test version has expired. Please contact Alistair Gillespie to renew"
 End If

'Dim Q As Double
Dim Plitt_Array(48, 10)
Dim Plitt_New(48, 4)
Dim Plitt_Summary(57, 4)
Dim Plitt_UF(9, 0)
Dim Plitt_OF(4, 0)
Dim I As Integer
Dim FeedMass As Double
Dim UFMass As Double
Dim OFMass As Double
Dim S As Double 'Liquor recovery to underflow
Dim Feed_Flow_Diluted As Double
Dim Feed_solids_Diluted As Double
Dim Flow_per_cyclone As Double
Dim Feed_solids_Diluted_mass_flow As Double
Dim Cv As Double
Dim PSD_d63 As Double
Dim PSD_Sharpness As Double
Dim Rs As Double 'Rs=recovery of solid to underflow
Dim Rf As Double 'Rf = recovery of fluid phase to the underflow
Dim UF_PSD_d63 As Double 'Weibull d63 for UF
Dim UF_PSD_Sharpness As Double 'Weibull sharpness for UF
Dim UF_d50 As Double
Dim UF_minus53 As Double
Dim slope53 As Double
Dim intercept53 As Double
Dim start As Double
Dim d50_Slope As Double
Dim d50_Intercept As Double
Dim d50_UF As Double

'Calculate inputs to Feed-PSD Weibull calculator
'Model coefficients obtained by regressing d63 and sharpness obtained from Weibull fits of plant vol-PSDs. See Minitab workbook Q23 - ARG - Weibull simulation of PSD.mpx
PSD_d63 = 224.5 - (0.555 * FeedMinus45) - (1.209 * (100 - FeedPlus150))
PSD_Sharpness = 1.965 - (0.07865 * FeedMinus45) + (0.01548 * (100 - FeedPlus150))

'Calculate Process parameters
Feed_Flow_Diluted = Feed_Flow + Dilution_Flow 'Total volumetric flow to cyclone cluster (m3/hr)
Feed_solids_Diluted = (Feed_Flow * Feed_Solids) / Feed_Flow_Diluted 'Solids concentration at cyclone inlet
Flow_per_cyclone = Feed_Flow_Diluted / No_cyclones ' Flow per cyclone (m3/hr)
Q = Flow_per_cyclone * 1000 / 60 ' Flow per cyclone (L/min)
Feed_solids_Diluted_mass_flow = Feed_Flow * Feed_Solids / 1000 ' Feed solids mass flow (t/hr)
Cv = Feed_solids_Diluted / rho_s / 10 'feed vol fraction solids as %

'Calculate Plitt factors
d50 = F * (50.5 * Dc ^ 0.46 * di ^ 0.6 * D_o ^ 1.21 * Exp(0.063 * Cv)) / (Du ^ 0.71 * h ^ 0.38 * Q ^ 0.45 * (rho_s - rho_l) ^ 0.5) 'd50 of grade efficiency curve
dP = PF * (1.88 * Q ^ 1.78 * Exp(0.0055 * Cv)) / (Dc ^ 0.37 * di ^ 0.94 * h ^ 0.28 * (Du ^ 2 + D_o ^ 2) ^ 0.87) '(8) Pressure Drop across the Cyclone in kPa
head = dP / (9.81 * rho_f) ' '(9) Pressure drop across cyclone, in metres of feed slurry
S = SFactor * (1.9 * (Du / D_o) ^ 3.31 * h ^ 0.54 * (Du ^ 2 + D_o ^ 2) ^ 0.36 * Exp(0.0054 * Cv)) / (head ^ 0.24 * Dc ^ 1.11) ' "S" value in recovery calc Rv = S / (S + 1) = (vol.flow rate in UF / vol.flow rate in OF)
Rv = S / (S + 1) '(10) Recovery of feed volume to the underflow product
m = SharpFactor * 1.94 * Exp((-1.58 * Rv) * (Dc ^ 2 * h / Q) ^ 0.15) 'Sharpness of grade efficiency curve

'Generate Plitt initial array:
'Column-0 = Diameter
'Column-1 = Rosin-Ramler grade efficiency
'Column 2 = Feed cumulative PSD (Weibull)
'Column-3 = Feed differential PSD (Weibull
Plitt_Array(0, 0) = 1
Plitt_Array(0, 1) = (1 - Exp(-0.693 * (Plitt_Array(0, 0) / d50) ^ m)) + Rv * (1 - (1 - Exp(-0.693 * (Plitt_Array(0, 0) / d50) ^ m))) 'Grade efficiency (Rosin-Rammler)
For I = 1 To 47
    Plitt_Array(I, 0) = I * 5 'Particle diameters
    Plitt_Array(I, 1) = (1 - Exp(-0.693 * (Plitt_Array(I, 0) / d50) ^ m)) + Rv * (1 - (1 - Exp(-0.693 * (Plitt_Array(I, 0) / d50) ^ m))) 'Grade efficiency (Rosin-Rammler)
    Plitt_Array(I, 2) = 1 - Exp(-(Plitt_Array(I, 0) / PSD_d63) ^ PSD_Sharpness)    'Cumulative feed PSD (Weibull)
    Plitt_Array(I, 3) = PSD_Sharpness / (PSD_d63 ^ PSD_Sharpness) * Plitt_Array(I, 0) ^ (PSD_Sharpness - 1) * Exp(-(Plitt_Array(I, 0) / PSD_d63) ^ PSD_Sharpness) 'Differential feed PSD (Weibull)
Next I

'Calculate Mass summation for normalised differential feed PSD (Weibull)
For I = 0 To 47
    FeedMass = FeedMass + Plitt_Array(I, 3)
Next I

'Calculate feed differential PSD (normalised), underflow  differential PSD (not normalised), overflow differential PSD (not normalised)
For I = 0 To 47
    Plitt_Array(I, 4) = Plitt_Array(I, 3) / FeedMass
    Plitt_Array(I, 5) = Plitt_Array(I, 1) * Plitt_Array(I, 4)
    Plitt_Array(I, 6) = (1 - Plitt_Array(I, 1)) * Plitt_Array(I, 4)
Next I

'Calculate underflow and overflow differential PSD (not normalised).
For I = 0 To 47
    UFMass = UFMass + Plitt_Array(I, 5)
    OFMass = OFMass + Plitt_Array(I, 6)
Next I

For I = 0 To 47
    Plitt_Array(I, 7) = Plitt_Array(I, 5) / UFMass
    Plitt_Array(I, 8) = Plitt_Array(I, 6) / OFMass
Next I

'Calculate underflow and overflow differential PSD (normalised).
Plitt_Array(I, 9) = Plitt_Array(I, 7)
Plitt_Array(I, 10) = Plitt_Array(I, 8)

'Calculate underflow and overflow cumulative PSD (normalised)
For I = 1 To 47
    Plitt_Array(I, 9) = Plitt_Array(I, 7) + Plitt_Array(I - 1, 9)
    Plitt_Array(I, 10) = Plitt_Array(I, 8) + Plitt_Array(I - 1, 10)
Next

'Calculate recoveries
Rs = Plitt_Array(47, 5) * 1000 'Solids recovery to underflow
Rf = Rv - (Cv / 100 * Rs) / (1 - (Cv / 100)) 'Fluid recovery to underflow

'Build output summary arrays
'PSD and grade efficiency table

    Plitt_Summary(0, 0) = "---Plitt Output---"
    Plitt_Summary(0, 1) = "-"
    Plitt_Summary(0, 2) = "-"
    Plitt_Summary(0, 3) = "-"
    Plitt_Summary(0, 4) = "-"
    
    Plitt_Summary(1, 0) = "-45 (%)   [feed : uf : of]" 'Diameter 45
    Plitt_Summary(1, 1) = Plitt_Array(9, 2) * 100 'Feed-45
    Plitt_Summary(1, 2) = Plitt_Array(9, 9) * 100 'UF-45
    Plitt_Summary(1, 3) = Plitt_Array(9, 10) * 100 'OF+150
    Plitt_Summary(1, 4) = "-"
    
    Plitt_Summary(2, 0) = "+150 (%)   [feed : uf : of]" 'Diameter 150
    Plitt_Summary(2, 1) = (1 - Plitt_Array(30, 2)) * 100 'Feed+150
    Plitt_Summary(2, 2) = (1 - Plitt_Array(30, 9)) * 100 'UF-45
    Plitt_Summary(2, 3) = (1 - Plitt_Array(30, 10)) * 100 'OF+150
    Plitt_Summary(2, 4) = "-"
     
    Plitt_Summary(3, 0) = "Flow/cyclone (m3/hr)"
    Plitt_Summary(3, 1) = Flow_per_cyclone
    Plitt_Summary(3, 2) = "-"
    Plitt_Summary(3, 3) = "-"
    Plitt_Summary(3, 4) = "-"
    
    Plitt_Summary(4, 0) = "Cyclone dP (kPa)"
    Plitt_Summary(4, 1) = dP
    Plitt_Summary(4, 2) = "-"
    Plitt_Summary(4, 3) = "-"
    Plitt_Summary(4, 4) = "-"
    
    Plitt_Summary(5, 0) = "Recovery (%) [solid : fluid : volume]"
    Plitt_Summary(5, 1) = Rs * 100
    Plitt_Summary(5, 2) = Rf * 100
    Plitt_Summary(5, 3) = Rv * 100
    Plitt_Summary(5, 4) = "-"
            
    Plitt_Summary(6, 0) = "Solid flow (t/h)  [feed : uf : of]"
    Plitt_Summary(6, 1) = Feed_solids_Diluted_mass_flow
    Plitt_Summary(6, 2) = Plitt_Summary(6, 1) * Plitt_Array(47, 5) * 1000
    Plitt_Summary(6, 3) = Plitt_Summary(6, 1) - Plitt_Summary(6, 2)
    Plitt_Summary(6, 4) = "-"
        
    Plitt_Summary(7, 0) = "Solids conc. (g/L)  [feed : uf : of]"
    Plitt_Summary(7, 1) = Feed_solids_Diluted
    Plitt_Summary(7, 2) = Plitt_Summary(6, 2) / Rv / Feed_Flow_Diluted * 1000
    Plitt_Summary(7, 3) = Plitt_Summary(6, 3) / (1 - Rv) / Feed_Flow_Diluted * 1000
    Plitt_Summary(7, 4) = "-"
    
    Plitt_Summary(8, 0) = "SPO (%)  [feed : uf : of]"
    Plitt_Summary(8, 1) = FeedSPO
    Plitt_Summary(8, 2) = FeedSPO * Feed_solids_Diluted_mass_flow / 100 * Rf / Plitt_Summary(6, 2) * 100
    Plitt_Summary(8, 3) = FeedSPO * Feed_solids_Diluted_mass_flow / 100 * (1 - Rf) / Plitt_Summary(6, 3) * 100
    Plitt_Summary(8, 4) = "-"
    
    Plitt_Summary(9, 0) = "<Diameter (um)>"
    Plitt_Summary(9, 1) = "<Feed>"
    Plitt_Summary(9, 2) = "<Underflow>"
    Plitt_Summary(9, 3) = "<Overflow>"
    Plitt_Summary(9, 4) = "<Grade>"
  
For I = 10 To 57
    Plitt_Summary(I, 0) = Plitt_Array(I - 10, 0) 'Diameters
    Plitt_Summary(I, 1) = Plitt_Array(I - 10, 2) * 100 'Cumulative feed PSD
    Plitt_Summary(I, 2) = Plitt_Array(I - 10, 9) * 100 'Cumulative underflow PSD
    Plitt_Summary(I, 3) = Plitt_Array(I - 10, 10) * 100 'Cumulative overflow PSD
    Plitt_Summary(I, 4) = Plitt_Array(I - 10, 1) * 100 'Grade efficiency curve
Next I

'Calculate Weibull parameters for underflow PSD
'The following doesn't work because the Weibull coeffs are for feed PSD only
'UF_PSD_d63 = 224.5 - (0.555 * Plitt_Array(19, 2)) - (1.209 * (100 - Plitt_Array(40, 2)))
'UF_PSD_Sharpness = 1.965 - (0.07865 * Plitt_Array(19, 2)) + (0.01548 * (100 - Plitt_Array(40, 2)))
'UF_minus53 = (1 - Exp(-53 / UF_PSD_d63) ^ UF_PSD_Sharpness) * 100

'Calculate underflow %-53 by linear interpolation between 50 and 55 micron
slope53 = (Plitt_Summary(21, 2) - Plitt_Summary(20, 2)) / (Plitt_Summary(21, 0) - Plitt_Summary(20, 0)) 'slope of UF PSD between 50 and 55 micron
intercept53 = Plitt_Summary(21, 2) - slope53 * Plitt_Summary(21, 0) 'intercept of UF PSD between 50 and 55 micron
UF_minus53 = 53 * slope53 + intercept53

'Calculate underflow d50 by linear interpolation of inverse PSD curve
'Assumes the UF PSD is relatively linear around d50; determines the first cumulative-PSD
'point above d50 (point N), then does a 2-pt linear interpolation between points N and N-1
For I = 10 To 57
    If Plitt_Summary(I, 2) > 50 Then 'find first point (point 'N') where cumulative vol is > 50%
        If Plitt_Summary(I - 1, 2) < 50 Then 'check if point 'N-1' has cumulative vol <50%
            d50_Slope = (Plitt_Summary(I - 1, 0) - Plitt_Summary(I, 0)) / (Plitt_Summary(I - 1, 2) - Plitt_Summary(I, 2)) ' calculate 2-pt slope
            d50_Intercept = Plitt_Summary(I - 1, 0) - d50_Slope * Plitt_Summary(I - 1, 2) 'calculate 2-pt intercept
            d50_UF = 50 * d50_Slope + d50_Intercept 'estimate d50
        End If
    End If
Next I

'Populate underflow array
Plitt_UF(0, 0) = Plitt_Summary(1, 2) 'UF -45
Plitt_UF(1, 0) = UF_minus53
Plitt_UF(2, 0) = Plitt_Summary(2, 2) 'UF +150
Plitt_UF(3, 0) = d50_UF
Plitt_UF(4, 0) = Plitt_Summary(6, 2) 'Solids mass flow
Plitt_UF(5, 0) = Plitt_Summary(7, 2) 'Solids concentration
Plitt_UF(6, 0) = Plitt_Summary(8, 2) 'Solids concentration
Plitt_UF(7, 0) = Rs * 100
Plitt_UF(8, 0) = Rf * 100
Plitt_UF(9, 0) = Rv * 100

'Populate overflow array
Plitt_OF(0, 0) = Plitt_Summary(1, 3) 'OF -45
Plitt_OF(1, 0) = Plitt_Summary(2, 3) 'OF +150
Plitt_OF(2, 0) = Plitt_Summary(6, 3) 'Solids mass flow
Plitt_OF(3, 0) = Plitt_Summary(7, 3) 'Solids concentration
Plitt_OF(4, 0) = Plitt_Summary(8, 3) 'Solids concentration

If Mode = 0 Then
Plitt = Plitt_Summary
ElseIf Mode = 1 Then
Plitt = Plitt_UF
ElseIf Mode = 2 Then
Plitt = Plitt_OF
ElseIf Mode = 3 Then
Plitt = dP
ElseIf Mode = 4 Then
Plitt = UF_minus53
ElseIf Mode = 5 Then
Plitt = d50_UF
End If
End Function
