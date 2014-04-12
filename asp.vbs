Class clsServer
	Public Function CreateObject(className) 
		Set CreateObject = CreateObject(className)
	End Sub
	
	Public Function HTMLEncode(value) 
		HTMLEncode = value
	End Function
End Class

Class clsResponse
	Public Property Let Buffer(value)
		
	End Property
	
	Public Sub Write(value)
		Wscript.Echo value
	End Sub
End Class

Class clsRequest	
	Public Function QueryString(key)
	
	End Function
	
	Public Function ServerVariables(key)
	
	End Function
	
	Public Function Form(key)
	
	End Function
	
	Public Function QueryString(key)
	
	End Function	
End Class

Class clsSession
	Public Property Let LCID(value)
	
	End Property
	
  Public Default Property Get Item( Key )
        
  End Property	
End class

Set Session = new clsSession
Set Server = new clsServer
Set Response = new clsResponse
Set Request = new clsRequest

