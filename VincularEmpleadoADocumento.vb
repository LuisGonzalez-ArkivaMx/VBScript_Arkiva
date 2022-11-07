' Luis Gonzalez - 15/06/2021 - Establecer el Nombre del Empleado en el Documento mediante una busqueda por CURP en el objeto Empleado

Option Explicit

' Inicializar Clases y Definicion de Propiedades
Dim IncapacidadesImss : IncapacidadesImss = Vault.ClassOperations.GetObjectClassIDByAlias("Class.IncapacidadesIMSS")
Dim SolicitudDeVacaciones : SolicitudDeVacaciones = Vault.ClassOperations.GetObjectClassIDByAlias("Class.SolicitudDeVacaciones")
Dim ServiceObservations : ServiceObservations = Vault.ClassOperations.GetObjectClassIDByAlias("Class.ServiceObservations")
Dim VoiceOfTheCustomer : VoiceOfTheCustomer = Vault.ClassOperations.GetObjectClassIDByAlias("Class.VoiceOfTheCustomer")
Dim PeakCertificado : PeakCertificado = Vault.ClassOperations.GetObjectClassIDByAlias("Class.PeakCertificado")
Dim ActasDisciplinarias : ActasDisciplinarias = Vault.ClassOperations.GetObjectClassIDByAlias("Class.ActasDisciplinarias")
Dim Reembolsos : Reembolsos = Vault.ClassOperations.GetObjectClassIDByAlias("Class.Reembolsos")
Dim CertificadosDeCursos : CertificadosDeCursos = Vault.ClassOperations.GetObjectClassIDByAlias("Class.CertificadosDeCursos")
Dim RecibosDeNomina : RecibosDeNomina = Vault.ClassOperations.GetObjectClassIDByAlias("Class.RecibosDeNomina")
Dim Empleado : Empleado = Vault.ClassOperations.GetObjectClassIDByAlias("Class.Empleado")
Dim propCurp : propCurp = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Curp")
Dim propEmpleado : propEmpleado = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Empleado")

' Inicializar Objetos M-Files
Dim oPropertyValues : Set oPropertyValues = CreateObject( "MFilesAPI.PropertyValues" )
Set oPropertyValues = Vault.ObjectPropertyOperations.GetProperties( ObjVer )
Dim oPropertyValue : Set oPropertyValue = CreateObject( "MFilesAPI.PropertyValue" )
Dim iClass : iClass = oPropertyValues.SearchForProperty( MFBuiltInPropertyDefClass ).TypedValue.GetLookupID

Dim szCurp, oResults, oResult

' Buscar el Empleado por medio de su CURP 
If iClass = IncapacidadesImss Or iClass = SolicitudDeVacaciones Or iClass = ServiceObservations Or _
	iClass = VoiceOfTheCustomer Or iClass = PeakCertificado Or iClass = ActasDisciplinarias Or _
	iClass = Reembolsos Or iClass = CertificadosDeCursos Or iClass = RecibosDeNomina Then
	
	szCurp = oPropertyValues.SearchForPropertyEx( propCurp, true ).TypedValue.GetValueAsLookup().DisplayValue
		
	If oPropertyValues.IndexOf( propEmpleado ) <> -1 Then
		
		If oPropertyValues.SearchForPropertyEx( propEmpleado, true ).TypedValue.IsNULL() Then
			
			Dim oSCs : Set oSCs = CreateObject( "MFilesAPI.SearchConditions" )
            Dim oSC : Set oSC = CreateObject( "MFilesAPI.SearchCondition" )
			Dim oLookups : Set oLookups = CreateObject( "MFilesAPI.Lookups" )
			Dim oLookup : Set oLookup = CreateObject( "MFilesAPI.Lookup" )

            'Crea los filtros de búsqueda
            'Clase
            oSC.ConditionType = MFConditionTypeEqual
            oSC.Expression.DataPropertyValuePropertyDef = MFBuiltInPropertyDefClass
            oSC.TypedValue.SetValue MFDatatypeLookup, Empleado
            oSCs.Add -1, oSC

            'Propiedad
            oSC.ConditionType = MFConditionTypeEqual
            oSC.Expression.DataPropertyValuePropertyDef = propCurp
            oSC.TypedValue.SetValue MFDatatypeText, szCurp ' MFDatatypeLookup
            oSCs.Add -1, oSC

            'No eliminado
            oSc.ConditionType = MFConditionTypeEqual
            oSc.Expression.DataStatusValueType = MFStatusTypeDeleted
            oSc.TypedValue.SetValue MFDatatypeBoolean, False
            oScs.Add -1, oSc
            
            'Ejecuta la busqueda
            Set oResults = Vault.ObjectSearchOperations.SearchForObjectsByConditions( oScs, MFSearchFlagNone, False )
			
			'Si existe el empleado se vincula al documento
			If oResults.Count > 0 Then
				
				'Err.Raise MFScriptCancel, "La CURP es: " + oResults.Count				
				For Each oResult In oResults
				
					oLookup.Item = oResult.ObjVer.ID
					oLookups.Add -1, oLookup
					
				Next
				
				oPropertyValue.PropertyDef = propEmpleado
				oPropertyValue.TypedValue.SetValueToMultiSelectLookup oLookups
				Vault.ObjectPropertyOperations.SetProperty ObjVer, oPropertyValue
				
			End If
			
		End If
		
	End If
	
End If
