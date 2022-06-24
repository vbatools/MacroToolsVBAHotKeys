﻿;~ Credit for this file mostly goes to jethrow: https://github.com/jethrow/UIA_Interface/blob/master/UIA_Interface.ahk
;~ I have added a lot of modifications to it, including custom methods for elements (eg element.Click()), UIA_Enum class etc

;~ UI Automation Constants: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671207(v=vs.85).aspx
;~ UI Automation Enumerations: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671210(v=vs.85).aspx
;~ http://www.autohotkey.com/board/topic/94619-ahk-l-screen-reader-a-tool-to-get-text-anywhere/

/* Questions:
	- better way to do __properties?
	- better way to do __properties for multiple versions of objects? (eg UIA_Element2 being and UIA_Element.__properties + UIA_Element2.__properties)
	- if method returns a SafeArray, should we return a Wrapped SafeArray, Raw SafeArray, or AHK Array. Currently we return wrapped AHK arrays for SafeArrays. Although SafeArrays are more convenient to loop over, this causes more confusion in users who are not familiar with SafeArrays (questions such as why are they 0-indexed not 1-indexed, why doesnt for k, v in SafeArray work properly etc). 
	- on UIA Interface conversion methods, how should the data be returned? wrapped/extracted or raw? should raw data be a ByRef param?
	- do variants need cleared? what about SysAllocString BSTRs?
	- do RECT struts need destroyed?
	- if returning wrapped data & raw is ByRef, will the wrapped data being released destroy the raw data?
	- returning varaint data other than vt=3|8|9|13|0x2000
	- Cached Members?
	- UIA Element existance - dependent on window being visible (non minimized)?
	- function(params, ByRef out="……")
	- better way of supporting differing versions of IUIAutomation (version 2, 3, 4)
	- Get methods vs property getter: currently we use properties when the item stores data, fetching the data is "cheap" and when it doesn't have side-effects, and in computationally expensive cases use Get...(). 
*/

class UIA_Base {
	__New(p="", flag=0, version="") {
		ObjInsert(this,"__Type","IUIAutomation" SubStr(this.__Class,5))
		,ObjInsert(this,"__Value",p)
		,ObjInsert(this,"__Flag",flag)
		,ObjInsert(this,"__Version",version)
	}
	__Get(member) {
		if member not in base,__UIA,TreeWalkerTrue,TrueCondition ; base & __UIA should act as normal
		{
			if raw:=SubStr(member,0)="*" ; return raw data - user should know what they are doing
				member:=SubStr(member,1,-1)
			if RegExMatch(this.__properties, "im)^" member ",(\d+),(\w+)", m) { ; if the member is in the properties. if not - give error message
				if (m2="VARIANT")	; return VARIANT data - DllCall output param different
					return UIA_Hr(DllCall(this.__Vt(m1), "ptr",this.__Value, "ptr",UIA_Variant(out)))? (raw?out:UIA_VariantData(out)):
				else if (m2="RECT") ; return RECT struct - DllCall output param different
					return UIA_Hr(DllCall(this.__Vt(m1), "ptr",this.__Value, "ptr",&(rect,VarSetCapacity(rect,16))))? (raw?out:UIA_RectToObject(rect)):
				else if (m2="double")
					return UIA_Hr(DllCall(this.__Vt(m1), "ptr",this.__Value, "Double*",out))?out:
				else if UIA_Hr(DllCall(this.__Vt(m1), "ptr",this.__Value, "ptr*",out))
					return raw?out:m2="BSTR"?StrGet(out):RegExMatch(m2,"i)IUIAutomation\K\w+",n)?(IsFunc(n)?UIA_%n%(out):new UIA_%n%(out)):out ; Bool, int, DWORD, HWND, CONTROLTYPEID, OrientationType? ifIUIAutomation___ is a function, that will be called first, if not then an object is created with the name
			} else if ObjHasKey(UIA_Enum, member) {
				return UIA_Enum[member]
			} else if RegexMatch(member, "i)PatternId|EventId|PropertyId|AttributeId|ControlTypeId|AnnotationType|StyleId|LandmarkTypeId|HeadingLevel|ChangeId|MetadataId", match) {
				return UIA_Enum["UIA_" match](member)
			} else throw Exception("Property not supported by the " this.__Class " Class.",-1,member)
		}
	}
	__Set(member) {
		if !(member == "base")
			throw Exception("Assigning values not supported by the " this.__Class " Class.",-1,member)
	}
	__Call(member, params*) {
		
		if RegexMatch(member, "i)^(?:UIA_)?(PatternId|EventId|PropertyId|AttributeId|ControlTypeId|AnnotationType|StyleId|LandmarkTypeId|HeadingLevel|ChangeId|MetadataId)$", match) {
			return UIA_Enum["UIA_" match1](params*)
		} else if !ObjHasKey(UIA_Base,member)&&!ObjHasKey(this,member)&&!"_NewEnum"
			throw Exception("Method Call not supported by the " this.__Class " Class.",-1,member)
	}
	__Delete() {
		this.__Flag ? ObjRelease(this.__Value):
	}
	__Vt(n) {
		return NumGet(NumGet(this.__Value+0,"ptr")+n*A_PtrSize,"ptr")
	}
}	

class UIA_Interface extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671406(v=vs.85).aspx
	static __IID := "{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}"
		,  __properties := "ControlViewWalker,14,IUIAutomationTreeWalker`r`nContentViewWalker,15,IUIAutomationTreeWalker`r`nRawViewWalker,16,IUIAutomationTreeWalker`r`nRawViewCondition,17,IUIAutomationCondition`r`nControlViewCondition,18,IUIAutomationCondition`r`nContentViewCondition,19,IUIAutomationCondition`r`nProxyFactoryMapping,48,IUIAutomationProxyFactoryMapping`r`nReservedNotSupportedValue,54,IUnknown`r`nReservedMixedAttributeValue,55,IUnknown"
	
	CompareElements(e1,e2) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr",e1.__Value, "ptr",e2.__Value, "int*",out))? out:
	}
	CompareRuntimeIds(r1,r2) {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "ptr",ComObjValue(r1), "ptr",ComObjValue(r2), "int*",out))? out:
	}
	GetRootElement() {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr*",out))? UIA_Element(out):
	}
	ElementFromHandle(hwnd, activateChromiumAccessibility=False) {
		if activateChromiumAccessibility {
			WinGet, cList, ControlList, ahk_id %hwnd%
			if InStr(cList, "Chrome_RenderWidgetHostHWND1")
				SendMessage, WM_GETOBJECT := 0x003D, 0, 1, Chrome_RenderWidgetHostHWND1, ahk_id %hwnd%
		}
		return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "ptr",hwnd, "ptr*",out))? UIA_Element(out):
	}
	ElementFromPoint(x="", y="", activateChromiumAccessibility=False) {
		if (x==""||y=="")
			DllCall("GetCursorPos","Int64*",pt)
		if (activateChromiumAccessibility && (hwnd := DllCall("GetAncestor", "UInt", DllCall("WindowFromPoint", pt), "UInt", GA_ROOT := 2))) { ; hwnd from point by SKAN
			WinGet, cList, ControlList, ahk_id %hwnd%
			if InStr(cList, "Chrome_RenderWidgetHostHWND1")
				SendMessage, WM_GETOBJECT := 0x003D, 0, 1, Chrome_RenderWidgetHostHWND1, ahk_id %hwnd%
		}
		return UIA_Hr(DllCall(this.__Vt(7), "ptr",this.__Value, "UInt64",x==""||y==""?pt:x&0xFFFFFFFF|(y&0xFFFFFFFF)<<32, "ptr*",out))? UIA_Element(out):
	}	
	GetFocusedElement() {
		return UIA_Hr(DllCall(this.__Vt(8), "ptr",this.__Value, "ptr*",out))? UIA_Element(out):
	}

	GetRootElementBuildCache(cacheRequest) { ; UNTESTED
		return UIA_Hr(DllCall(this.__Vt(9), "ptr",this.__Value, "ptr", cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}
	ElementFromHandleBuildCache(hwnd, cacheRequest) {
		return UIA_Hr(DllCall(this.__Vt(10), "ptr",this.__Value, "ptr",hwnd, "ptr",cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}

	ElementFromPointBuildCache(x="", y="", cacheRequest="") { ; UNTESTED
		return UIA_Hr(DllCall(this.__Vt(11), "ptr",this.__Value, "UInt64",x==""||y==""?0*DllCall("GetCursorPos","Int64*",pt)+pt:x&0xFFFFFFFF|(y&0xFFFFFFFF)<<32, "ptr", cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}	

	GetFocusedElementBuildCache(cacheRequest) { ; UNTESTED
		return UIA_Hr(DllCall(this.__Vt(12), "ptr",this.__Value, "ptr", cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}
	CreateTreeWalker(condition) {
		return UIA_Hr(DllCall(this.__Vt(13), "ptr",this.__Value, "ptr",Condition.__Value, "ptr*",out))? new UIA_TreeWalker(out):
	}
	
	CreateCacheRequest() {
		return UIA_Hr(DllCall(this.__Vt(20), "ptr",this.__Value, "ptr*",out))? new UIA_CacheRequest(out):
	}

	CreateTrueCondition() {
		return UIA_Hr(DllCall(this.__Vt(21), "ptr",this.__Value, "ptr*",out))? new UIA_BoolCondition(out):
	}
	CreateFalseCondition() {
		return UIA_Hr(DllCall(this.__Vt(22), "ptr",this.__Value, "ptr*",out))? new UIA_BoolCondition(out):
	}
	CreatePropertyCondition(propertyId, var, type="Variant") {
		if (type!="Variant")
			UIA_Variant(var,type,var)
		else if (maybeVar := UIA_Enum.UIA_PropertyVariantType(propertyId)) {
			UIA_Variant(var,maybeVar,var)
		}
		return UIA_Hr((A_PtrSize == 4) ? DllCall(this.__Vt(23), "ptr",this.__Value, "int",propertyId, "int64", NumGet(var, 0, "int64"), "int64", NumGet(var, 8, "int64"), "ptr*",out) : DllCall(this.__Vt(23), "ptr",this.__Value, "int",propertyId, "ptr",&var, "ptr*",out))? new UIA_PropertyCondition(out):
	}
	CreatePropertyConditionEx(propertyId, var, type="Variant", flags=0x1) {
		; PropertyConditionFlags_IgnoreCase = 0x1
		if (type!="Variant")
			UIA_Variant(var,type,var)
		else if (maybeVar := UIA_Enum.UIA_PropertyVariantType(propertyId)) {
			UIA_Variant(var,maybeVar,var)
		}
		return UIA_Hr((A_PtrSize == 4) ? DllCall(this.__Vt(24), "ptr",this.__Value, "int",propertyId, "int64", NumGet(var, 0, "int64"), "int64", NumGet(var, 8, "int64"), "uint",flags, "ptr*",out) : DllCall(this.__Vt(24), "ptr",this.__Value, "int",propertyId, "ptr",&var, "uint",flags, "ptr*",out))? new UIA_PropertyCondition(out):
	}
	CreateAndCondition(c1,c2) {
		return UIA_Hr(DllCall(this.__Vt(25), "ptr",this.__Value, "ptr",c1.__Value, "ptr",c2.__Value, "ptr*",out))? new UIA_AndCondition(out):
	}
	CreateAndConditionFromArray(array) { ; ComObj(0x2003)??
	;->in: AHK Array or Wrapped SafeArray
		if ComObjValue(array)&0x2000
			SafeArray:=array
		else {
			SafeArray:=ComObj(0x2003,DllCall("oleaut32\SafeArrayCreateVector", "uint",13, "uint",0, "uint",array.MaxIndex()),1)
			for i,c in array
				SafeArray[A_Index-1]:=c.__Value, ObjAddRef(c.__Value) ; AddRef - SafeArrayDestroy will release UIA_Conditions - they also release themselves
		}
		return UIA_Hr(DllCall(this.__Vt(26), "ptr",this.__Value, "ptr",ComObjValue(SafeArray), "ptr*",out))? new UIA_AndCondition(out):
	}
	CreateAndConditionFromNativeArray(conditions, conditionCount) { ; UNTESTED
		/*	[in]           IUIAutomationCondition **conditions,
			[in]           int conditionCount,
			[out, retval]  IUIAutomationCondition **newCondition
		*/
		return UIA_Hr(DllCall(this.__Vt(27), "ptr",this.__Value, "ptr", conditions, "int", conditionCount, "ptr*",out))? new UIA_AndCondition(out):
	}
	CreateOrCondition(c1,c2) {
		return UIA_Hr(DllCall(this.__Vt(28), "ptr",this.__Value, "ptr",c1.__Value, "ptr",c2.__Value, "ptr*",out))? new UIA_OrCondition(out):
	}
	CreateOrConditionFromArray(array) {
	;->in: AHK Array or Wrapped SafeArray
		if ComObjValue(array)&0x2000 ; equivalent to ComObjType(array) == VT_ARRAY ?
			SafeArray:=array
		else {
			SafeArray:=ComObj(0x2003,DllCall("oleaut32\SafeArrayCreateVector", "uint",13, "uint",0, "uint",array.MaxIndex()),1)
			for i,c in array
				SafeArray[A_Index-1]:=c.__Value, ObjAddRef(c.__Value) ; AddRef - SafeArrayDestroy will release UIA_Conditions - they also release themselves
		}
		return UIA_Hr(DllCall(this.__Vt(29), "ptr",this.__Value, "ptr",ComObjValue(SafeArray), "ptr*",out))? new UIA_OrCondition(out):
	}
	CreateOrConditionFromNativeArray(p*) { ; Not Implemented
		return UIA_Hr(DllCall(this.__Vt(30), "ptr",this.__Value, "ptr",conditions, "int", conditionCount, "ptr*",out))? new UIA_OrCondition(out):
	/*	[in]           IUIAutomationCondition **conditions,
		[in]           int conditionCount,
		[out, retval]  IUIAutomationCondition **newCondition
	*/
	}
	CreateNotCondition(c) {
		return UIA_Hr(DllCall(this.__Vt(31), "ptr",this.__Value, "ptr",c.__Value, "ptr*",out))? new UIA_NotCondition(out):
	}
	AddAutomationEventHandler(eventId, element, scope=0x4, cacheRequest=0, handler="") {
		return UIA_Hr(DllCall(this.__Vt(32), "ptr",this.__Value, "int", eventId, "ptr", element.__Value, "uint", scope, "ptr",cacheRequest.__Value,"ptr",handler.__Value))
	}
	RemoveAutomationEventHandler(eventId, element, handler) {
		return UIA_Hr(DllCall(this.__Vt(33), "ptr",this.__Value, "int", eventId, "ptr", element.__Value, "ptr",handler.__Value))
	}

	;~ AddPropertyChangedEventHandlerNativeArray 	34
	
	AddPropertyChangedEventHandler(element,scope=0x1,cacheRequest=0,handler="",propertyArray="") {
		SafeArray:=ComObjArray(0x3,propertyArray.MaxIndex())
		for i,propertyId in propertyArray
			SafeArray[i-1]:=propertyId
		return UIA_Hr(DllCall(this.__Vt(35), "ptr",this.__Value, "ptr",element.__Value, "int",scope, "ptr",cacheRequest.__Value,"ptr",handler.__Value,"ptr",ComObjValue(SafeArray)))
	}
	RemovePropertyChangedEventHandler(element, handler) {
		return UIA_Hr(DllCall(this.__Vt(36), "ptr",this.__Value, "ptr",element.__Value, "ptr", handler.__Value))
	}

	AddStructureChangedEventHandler(element, handler) { ; UNTESTED
		return UIA_Hr(DllCall(this.__Vt(37), "ptr",this.__Value, "ptr",element.__Value, "ptr",handler.__Value))
	}
	RemoveStructureChangedEventHandler(element, handler) { ; UNTESTED
		return UIA_Hr(DllCall(this.__Vt(38), "ptr",this.__Value, "ptr", element.__Value, "ptr",handler.__Value))
	}

	AddFocusChangedEventHandler(handler="", cacheRequest=0) { ; handler is required, cacheRequest can be left to 0
		return UIA_Hr(DllCall(this.__Vt(39), "ptr",this.__Value, "ptr",cacheRequest.__Value, "ptr",handler.__Value))
	}
	RemoveFocusChangedEventHandler(handler) {
		return UIA_Hr(DllCall(this.__Vt(40), "ptr",this.__Value, "ptr",handler.__Value))
	}
	RemoveAllEventHandlers() {
		return UIA_Hr(DllCall(this.__Vt(41), "ptr",this.__Value))
	}

	IntNativeArrayToSafeArray(ByRef nArr, n="") {
		return UIA_Hr(DllCall(this.__Vt(42), "ptr",this.__Value, "ptr",&nArr, "int",n?n:VarSetCapacity(nArr)/4, "ptr*",out))? ComObj(0x2003,out,1):
	}
	IntSafeArrayToNativeArray(sArr, Byref nArr, Byref arrayCount) { ; NOT WORKING
		VarSetCapacity(nArr,(sArr.MaxIndex()+1)*A_PtrSize)
		return UIA_Hr(DllCall(this.__Vt(43), "ptr",this.__Value, "ptr",ComObjValue(sArr), "ptr*",nArr, "int*",arrayCount))? nArr:
	}

	RectToVariant(ByRef rect, ByRef out="") {	; in:{left,top,right,bottom} ; out:(left,top,width,height)
		; in:	RECT Struct
		; out:	AHK Wrapped SafeArray & ByRef Variant
		return UIA_Hr(DllCall(this.__Vt(44), "ptr",this.__Value, "ptr",&rect, "ptr",UIA_Variant(out)))? UIA_VariantData(out):
	}
	VariantToRect(ByRef var, ByRef rect="") { ; NOT WORKING
		; in:	VT_VARIANT (SafeArray)
		; out:	AHK Wrapped RECT Struct & ByRef Struct
		VarSetCapacity(rect,16)
		return UIA_Hr(DllCall(this.__Vt(45), "ptr",this.__Value, "ptr",var, "ptr*",rect))? UIA_RectToObject(rect):
	}

	;~ SafeArrayToRectNativeArray 	46
	;~ CreateProxyFactoryEntry 	47
	GetPropertyProgrammaticName(Id) {
		return UIA_Hr(DllCall(this.__Vt(49), "ptr",this.__Value, "int",Id, "ptr*",out))? StrGet(out):
	}
	GetPatternProgrammaticName(Id) {
		return UIA_Hr(DllCall(this.__Vt(50), "ptr",this.__Value, "int",Id, "ptr*",out))? StrGet(out):
	}
	PollForPotentialSupportedPatterns(e, Byref Ids="", Byref Names="") { ; Returns an object where keys are the names and values are the Ids
		return UIA_Hr(DllCall(this.__Vt(51), "ptr",this.__Value, "ptr",e.__Value, "ptr*",Ids, "ptr*",Names))? UIA_SafeArraysToObject(Names:=ComObj(0x2008,Names,1),Ids:=ComObj(0x2003,Ids,1)):
	}
	PollForPotentialSupportedProperties(e, Byref Ids="", Byref Names="") {
		return UIA_Hr(DllCall(this.__Vt(52), "ptr",this.__Value, "ptr",e.__Value, "ptr*",Ids, "ptr*",Names))? UIA_SafeArraysToObject(Names:=ComObj(0x2008,Names,1),Ids:=ComObj(0x2003,Ids,1)):
	}
	CheckNotSupported(value) { ; Useless in this Framework???
	/*	Checks a provided VARIANT to see if it contains the Not Supported identifier.
		After retrieving a property for a UI Automation element, call this method to determine whether the element supports the 
		retrieved property. CheckNotSupported is typically called after calling a property retrieving method such as GetCurrentPropertyValue.
	*/
		return UIA_Hr(DllCall(this.__Vt(53), "ptr",this.__Value, "ptr",value, "int*",out))? out:
	}
	
	;~ ReservedNotSupportedValue 	54
	;~ ReservedMixedAttributeValue 	55
	
	ElementFromIAccessible(IAcc, childId=0) {
	/* The method returns E_INVALIDARG - "One or more arguments are not valid" - if the underlying implementation of the
	Microsoft UI Automation element is not a native Microsoft Active Accessibility server; that is, if a client attempts to retrieve
	the IAccessible interface for an element originally supported by a proxy object from Oleacc.dll, or by the UIA-to-MSAA Bridge.
	*/
		return UIA_Hr(DllCall(this.__Vt(56), "ptr",this.__Value, "ptr",ComObjValue(IAcc), "int",childId, "ptr*",out))? UIA_Element(out):
	}
	ElementFromIAccessibleBuildCache(IAcc, childId, cacheRequest) {
		return UIA_Hr(DllCall(this.__Vt(57), "ptr",this.__Value, "ptr",ComObjValue(IAcc), "int",childId, "ptr", cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}

	; ------- ONLY CUSTOM FUNCTIONS FROM HERE ON ----------------
	CreateCondition(property, val, flags=0) { ; flags: 0=no flags; 1=ignore case; 2=match substring; 3=ignore case and match substring
		if !RegexMatch(property, "^\d+$")
			RegexMatch(property, "(?:UIA_)?\K.+?(?=(Id)?PropertyId|$)", property), propCond := UIA_Enum.UIA_PropertyId(property), property := StrReplace(StrReplace(property, "AnnotationAnnotation", "Annotation"), "StylesStyle", "Style")
		else
			propCond := property
		if RegexMatch(val, "^\w+$") {
			val := IsFunc("UIA_Enum.UIA_" property "Id") ? UIA_Enum["UIA_" property "Id"](val) : IsFunc("UIA_Enum.UIA_" property) ? UIA_Enum["UIA_" property](val) : val
		}
		if (propCond && val)
			return this.CreatePropertyConditionEx(propCond, val,, flags)
	}
}

class UIA_Interface2 extends UIA_Interface {
	static __IID := "{34723aff-0c9d-49d0-9896-7ab52df8cd8a}"
		, __Properties := UIA_Interface.__Properties

	AutoSetFocus[] 
	{
		get {
			return UIA_Hr(DllCall(this.__Vt(58), "ptr",this.__Value, "ptr*", out))?out:
		}
		set {
			return UIA_Hr(DllCall(this.__Vt(59), "ptr",this.__Value, "int", value))
		}
	}
	ConnectionTimeout[] 
	{
		get {
			return UIA_Hr(DllCall(this.__Vt(60), "ptr",this.__Value, "ptr*", out))?out:
		}
		set {
			return UIA_Hr(DllCall(this.__Vt(61), "ptr",this.__Value, "int", value))
		}
	}
	TransactionTimeout[] 
	{
		get {
			return UIA_Hr(DllCall(this.__Vt(62), "ptr",this.__Value, "ptr*", out))?out:
		}
		set {
			return UIA_Hr(DllCall(this.__Vt(63), "ptr",this.__Value, "int", value))
		}
	}
}

class UIA_Interface3 extends UIA_Interface2 {
	static __IID := "{73d768da-9b51-4b89-936e-c209290973e7}"
		, __Properties := UIA_Interface2.__Properties
	AddTextEditTextChangedEventHandler(element, scope, textEditChangeType, cacheRequest=0, handler="") {
		return UIA_Hr(DllCall(this.__Vt(64), "ptr",this.__Value, "ptr", element.__Value, "int", scope, "int", textEditChangeType, "ptr", cacheRequest.__Value, "ptr", handler.__Value))
	}
	RemoveTextEditTextChangedEventHandler(element, handler) {
		return UIA_Hr(DllCall(this.__Vt(65), "ptr",this.__Value, "ptr", element.__Value, "ptr", handler.__Value))
	}
}

class UIA_Interface4 extends UIA_Interface3 {
	static __IID := "{1189c02a-05f8-4319-8e21-e817e3db2860}"
		, __Properties := UIA_Interface3.__Properties
	AddChangesEventHandler(element, scope, changeTypes, changesCount, cacheRequest=0, handler="") {
		return UIA_Hr(DllCall(this.__Vt(66), "ptr",this.__Value, "ptr", element.__Value, "int", scope, "int", changeTypes, "int", changesCount, "ptr", cacheRequest.__Value, "ptr", handler.__Value))
	}
	RemoveChangesEventHandler(element, handler) {
		return UIA_Hr(DllCall(this.__Vt(67), "ptr",this.__Value, "ptr", element.__Value, "ptr", handler.__Value))
	}
}
class UIA_Interface5 extends UIA_Interface4 {
	static __IID := "{25f700c8-d816-4057-a9dc-3cbdee77e256}"
		, __Properties := UIA_Interface4.__Properties
	AddNotificationEventHandler(element, scope, cacheRequest, handler) {
		return UIA_Hr(DllCall(this.__Vt(68), "ptr",this.__Value, "ptr", element.__Value, "int", scope, "ptr", cacheRequest.__Value, "ptr", handler.__Value))
	}
	RemoveNotificationEventHandler(element, handler) {
		return UIA_Hr(DllCall(this.__Vt(69), "ptr",this.__Value, "ptr", element.__Value, "ptr", handler.__Value))
	}
}
class UIA_Interface6 extends UIA_Interface5 { ; NOT IMPLEMENTED
	static __IID := "{aae072da-29e3-413d-87a7-192dbf81ed10}"
		, __Properties := UIA_Interface5.__Properties
/*
#define IUIAutomation6_CreateEventHandlerGroup(This,handlerGroup)	\
    ( (This)->lpVtbl -> CreateEventHandlerGroup(This,handlerGroup) ) 

#define IUIAutomation6_AddEventHandlerGroup(This,element,handlerGroup)	\
    ( (This)->lpVtbl -> AddEventHandlerGroup(This,element,handlerGroup) ) 

#define IUIAutomation6_RemoveEventHandlerGroup(This,element,handlerGroup)	\
    ( (This)->lpVtbl -> RemoveEventHandlerGroup(This,element,handlerGroup) ) 

#define IUIAutomation6_get_ConnectionRecoveryBehavior(This,connectionRecoveryBehaviorOptions)	\
    ( (This)->lpVtbl -> get_ConnectionRecoveryBehavior(This,connectionRecoveryBehaviorOptions) ) 

#define IUIAutomation6_put_ConnectionRecoveryBehavior(This,connectionRecoveryBehaviorOptions)	\
    ( (This)->lpVtbl -> put_ConnectionRecoveryBehavior(This,connectionRecoveryBehaviorOptions) ) 

#define IUIAutomation6_get_CoalesceEvents(This,coalesceEventsOptions)	\
    ( (This)->lpVtbl -> get_CoalesceEvents(This,coalesceEventsOptions) ) 

#define IUIAutomation6_put_CoalesceEvents(This,coalesceEventsOptions)	\
    ( (This)->lpVtbl -> put_CoalesceEvents(This,coalesceEventsOptions) ) 

#define IUIAutomation6_AddActiveTextPositionChangedEventHandler(This,element,scope,cacheRequest,handler)	\
    ( (This)->lpVtbl -> AddActiveTextPositionChangedEventHandler(This,element,scope,cacheRequest,handler) ) 

#define IUIAutomation6_RemoveActiveTextPositionChangedEventHandler(This,element,handler)	\
    ( (This)->lpVtbl -> RemoveActiveTextPositionChangedEventHandler(This,element,handler) ) 
*/
}
class UIA_Interface7 extends UIA_Interface6 {
	static __IID := "{29de312e-83c6-4309-8808-e8dfcb46c3c2}"
		, __Properties := UIA_Interface6.__Properties
}

class UIA_Element extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671425(v=vs.85).aspx
	static __IID := "{d22108aa-8ac5-49a5-837b-37bbb3d7591e}"
		,  __properties := "CurrentProcessId,20,int`r`nCurrentControlType,21,CONTROLTYPEID`r`nCurrentLocalizedControlType,22,BSTR`r`nCurrentName,23,BSTR`r`nCurrentAcceleratorKey,24,BSTR`r`nCurrentAccessKey,25,BSTR`r`nCurrentHasKeyboardFocus,26,BOOL`r`nCurrentIsKeyboardFocusable,27,BOOL`r`nCurrentIsEnabled,28,BOOL`r`nCurrentAutomationId,29,BSTR`r`nCurrentClassName,30,BSTR`r`nCurrentHelpText,31,BSTR`r`nCurrentCulture,32,int`r`nCurrentIsControlElement,33,BOOL`r`nCurrentIsContentElement,34,BOOL`r`nCurrentIsPassword,35,BOOL`r`nCurrentNativeWindowHandle,36,UIA_HWND`r`nCurrentItemType,37,BSTR`r`nCurrentIsOffscreen,38,BOOL`r`nCurrentOrientation,39,OrientationType`r`nCurrentFrameworkId,40,BSTR`r`nCurrentIsRequiredForForm,41,BOOL`r`nCurrentItemStatus,42,BSTR`r`nCurrentBoundingRectangle,43,RECT`r`nCurrentLabeledBy,44,IUIAutomationElement`r`nCurrentAriaRole,45,BSTR`r`nCurrentAriaProperties,46,BSTR`r`nCurrentIsDataValidForForm,47,BOOL`r`nCurrentControllerFor,48,IUIAutomationElementArray`r`nCurrentDescribedBy,49,IUIAutomationElementArray`r`nCurrentFlowsTo,50,IUIAutomationElementArray`r`nCurrentProviderDescription,51,BSTR`r`nCachedProcessId,52,int`r`nCachedControlType,53,CONTROLTYPEID`r`nCachedLocalizedControlType,54,BSTR`r`nCachedName,55,BSTR`r`nCachedAcceleratorKey,56,BSTR`r`nCachedAccessKey,57,BSTR`r`nCachedHasKeyboardFocus,58,BOOL`r`nCachedIsKeyboardFocusable,59,BOOL`r`nCachedIsEnabled,60,BOOL`r`nCachedAutomationId,61,BSTR`r`nCachedClassName,62,BSTR`r`nCachedHelpText,63,BSTR`r`nCachedCulture,64,int`r`nCachedIsControlElement,65,BOOL`r`nCachedIsContentElement,66,BOOL`r`nCachedIsPassword,67,BOOL`r`nCachedNativeWindowHandle,68,UIA_HWND`r`nCachedItemType,69,BSTR`r`nCachedIsOffscreen,70,BOOL`r`nCachedOrientation,71,OrientationType`r`nCachedFrameworkId,72,BSTR`r`nCachedIsRequiredForForm,73,BOOL`r`nCachedItemStatus,74,BSTR`r`nCachedBoundingRectangle,75,RECT`r`nCachedLabeledBy,76,IUIAutomationElement`r`nCachedAriaRole,77,BSTR`r`nCachedAriaProperties,78,BSTR`r`nCachedIsDataValidForForm,79,BOOL`r`nCachedControllerFor,80,IUIAutomationElementArray`r`nCachedDescribedBy,81,IUIAutomationElementArray`r`nCachedFlowsTo,82,IUIAutomationElementArray`r`nCachedProviderDescription,83,BSTR"

	SetFocus() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
	}
	GetRuntimeId(ByRef stringId="") {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "ptr*",sa))? UIA_SafeArrayToAHKArray(ComObj(0x2003,sa,1)):
	}
	FindFirst(c="", scope=0x4) {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "uint",scope, "ptr",(c=""?this.TrueCondition:c).__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	FindAll(c="", scope=0x4) {
		return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "uint",scope, "ptr",(c=""?this.TrueCondition:c).__Value, "ptr*",out))&&out? UIA_ElementArray(out):
	}
	FindFirstBuildCache(c="", scope=0x4, cacheRequest="") { ; UNTESTED
		return UIA_Hr(DllCall(this.__Vt(7), "ptr",this.__Value, "uint",scope, "ptr",(c=""?this.TrueCondition:c).__Value, "ptr",cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}
	FindAllBuildCache(c="", scope=0x4, cacheRequest="") { ; UNTESTED
		return UIA_Hr(DllCall(this.__Vt(8), "ptr",this.__Value, "uint",scope, "ptr",(c=""?this.TrueCondition:c).__Value, "ptr",cacheRequest.__Value, "ptr*",out))? UIA_ElementArray(out):
	}
	BuildUpdatedCache(cacheRequest) { ; UNTESTED
		return UIA_Hr(DllCall(this.__Vt(9), "ptr",this.__Value, "ptr", cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}
	GetCurrentPropertyValue(propertyId, ByRef out="") {
		if propertyId is not integer
			propertyId := UIA_Enum.UIA_PropertyId(propertyId)
		return UIA_Hr(DllCall(this.__Vt(10), "ptr",this.__Value, "uint", propertyId, "ptr",UIA_Variant(out)))? UIA_VariantData(out):
		
	}
	GetCurrentPropertyValueEx(propertyId, ignoreDefaultValue=1, ByRef out="") {
	; Passing FALSE in the ignoreDefaultValue parameter is equivalent to calling GetCurrentPropertyValue
		if propertyId is not integer
			propertyId := UIA_Enum.UIA_PropertyId(propertyId)
		return UIA_Hr(DllCall(this.__Vt(11), "ptr",this.__Value, "uint",propertyId, "uint",ignoreDefaultValue, "ptr",UIA_Variant(out)))? UIA_VariantData(out):
	}
	GetCachedPropertyValue(propertyId, ByRef out="") { ; UNTESTED
		if propertyId is not integer
			propertyId := UIA_Enum.UIA_PropertyId(propertyId)
		return UIA_Hr(DllCall(this.__Vt(12), "ptr",this.__Value, "uint",propertyId, "ptr",UIA_Variant(out)))? UIA_VariantData(out):
	}
	GetCachedPropertyValueEx(propertyId, ignoreDefaultValue=1, ByRef out="") {
		if propertyId is not integer
			propertyId := UIA_Enum.UIA_PropertyId(propertyId)
	; Passing FALSE in the ignoreDefaultValue parameter is equivalent to calling GetCachedPropertyValue
		return UIA_Hr(DllCall(this.__Vt(13), "ptr",this.__Value, "uint",propertyId, "uint",ignoreDefaultValue, "ptr",UIA_Variant(out)))? UIA_VariantData(out):
	}
	GetCurrentPatternAs(pattern="", ByRef usedPattern="") { ; If a full pattern name is specified then that exact version will be used (eg "TextPattern" will return a UIA_TextPattern object), otherwise the highest version will be used (eg "Text" might return UIA_TextPattern2 if it is available). usedPattern will be set to the actual string used to look for the pattern (used mostly for debugging purposes)
		if (InStr(usedPattern:=pattern, "Pattern")||(usedPattern := UIA_Pattern(pattern)))
			return UIA_Hr(DllCall(this.__Vt(14), "ptr",this.__Value, "int",UIA_%usedPattern%.__PatternId, "ptr",UIA_GUID(riid,UIA_%usedPattern%.__iid), "ptr*",out)) ? new UIA_%usedPattern%(out):
		else throw Exception("Pattern not implemented.",-1, "UIA_" pattern "Pattern")
	}
	;~ GetCachedPatternAs 	15	void **ppv
	GetCachedPatternAs(pattern="", ByRef usedPattern="") {
		if (InStr(usedPattern:=pattern, "Pattern")||(usedPattern := UIA_Pattern(pattern)))
			return UIA_Hr(DllCall(this.__Vt(15), "ptr",this.__Value, "int",UIA_%usedPattern%.__PatternId, "ptr",UIA_GUID(riid,UIA_%usedPattern%.__iid), "ptr*",out)) ? new UIA_%usedPattern%(out):
		else throw Exception("Pattern not implemented.",-1, "UIA_" pattern "Pattern")
	}
	;~ GetCurrentPattern 	16	Iunknown **patternObject
	GetCurrentPattern(patternId) {
		UIA_Hr(DllCall(this.__Vt(16), "ptr",this.__Value, "int",patternId, "ptr*",out))
		if name:=UIA_Enum.UIA_PatternId(patternId)
			return new UIA_%name%(out)
		return out
	}
	;~ GetCachedPattern 	17	Iunknown **patternObject
	GetCachedPattern(patternId) {
		UIA_Hr(DllCall(this.__Vt(17), "ptr",this.__Value, "int",patternId, "ptr*",out))
		if name:=UIA_Enum.UIA_PatternId(patternId)
			return new UIA_%name%(out)
		return out
	}
	;~ GetCachedParent 	18	IUIAutomationElement
	GetCachedParent() {
		return UIA_Hr(DllCall(this.__Vt(18), "ptr",this.__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	GetCachedChildren() { ; Haven't successfully tested
		return UIA_Hr(DllCall(this.__Vt(19), "ptr",this.__Value, "ptr*",out))&&out? UIA_ElementArray(out):
	}
	;~ GetClickablePoint 	84	POINT, BOOL
	GetClickablePoint() {
		UIA_Hr(DllCall(this.__Vt(84), "ptr",this.__Value, "ptr", &(point,VarSetCapacity(point,8)), "ptr*", out))&&out? {x:NumGet(point,0,"int"), y:NumGet(point,4,"int")}:
	}
	
	; ------- ONLY CUSTOM FUNCTIONS FROM HERE ON ----------------
	CurrentValue[] { ; Gets or sets the current value of the element. Getter is a wrapper for GetCurrentPropertyValue("Value"), setter a wrapper for SetValue
		get {
			return this.GetCurrentPropertyValue("Value")
		}
		set {
			return this.SetValue(value)
		}
	}
	GetClickablePointRelativeTo(relativeTo="") { ; relativeTo can be window, screen or client, default is A_CoordModeMouse
		res := this.GetClickablePoint()
		relativeTo := (relativeTo == "") ? A_CoordModeMouse : relativeTo
		StringLower, relativeTo, relativeTo
		if (relativeTo == "screen")
			return res
		else {
			hwnd := this.GetParentHwnd()
			if ((relativeTo == "window") || (relativeTo == "relative")) {
				VarSetCapacity(RECT, 16)
				DllCall("user32\GetWindowRect", "Ptr", hwnd, "Ptr", &RECT)
				return {x:(res.x-NumGet(&RECT, 0, "Int")), y:(res.y-NumGet(&RECT, 4, "Int"))}
			} else if (relativeTo == "client") {
				VarSetCapacity(pt,8,0), NumPut(res.x,&pt,0,"int"), NumPut(res.y,&pt,4,"int")
				DllCall("ScreenToClient", "Ptr",hwnd, "Ptr",&pt)
				return {x:NumGet(pt,"int"), y:NumGet(pt,4,"int")}
			}
		}
	}

	GetSupportedPatterns() { ; Get all available patterns for the element. Use of this should be avoided, since it calls GetCurrentPatternAs for every possible pattern. A better option is PollForPotentialSupportedPatterns.
		result := []
		patterns := "Invoke,Selection,Value,RangeValue,Scroll,ExpandCollapse,Grid,GridItem,MultipleView,Window,SelectionItem,Dock,Table,TableItem,Text,Toggle,Transform,ScrollItem,ItemContainer,VirtualizedItem,SyncronizedInput,LegacyIAccessible"

		Loop, Parse, patterns, `,
		{
			try {
				if this.GetCurrentPropertyValue(UIA_Enum.UIA_PropertyId("Is" A_LoopField "PatternAvailable")) {
					result.Push(A_LoopField)
				}
			}
		}
		return result
	}
	GetParentHwnd() { ; Get the parent window hwnd from the element
		hwndNotZeroCond := this.__UIA.CreateNotCondition(this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_PropertyId("NativeWindowHandle"), 0)) ; create a condition to find NativeWindowHandlePropertyId of not 0
		TW := this.__UIA.CreateTreeWalker(hwndNotZeroCond)
		try {
			hwnd := TW.NormalizeElement(this).GetCurrentPropertyValue(UIA_Enum.UIA_PropertyId("NativeWindowHandle"))
			return hwndRoot := DllCall("user32\GetAncestor", Ptr,hwnd, UInt,2, Ptr)
		} catch {
			return 0
		}
	}
	SetValue(val, pattern="") { ; Set element value using Value pattern, or as a fall-back using LegacyIAccessible pattern. If a pattern is specified then that is used instead.
		if !pattern {
			try {
				this.GetCurrentPatternAs("Value").SetValue(val)
			} catch {
				this.GetCurrentPatternAs("LegacyIAccessible").SetValue(val)
			}
		} else {
			this.GetCurrentPatternAs(pattern).SetValue(val)
		}
	}
	Click(but="", clickCount=1) { ; Click using LegacyIAccessible pattern DoDefaultAction(), or the Toggle pattern with Toggle() method or Invoke pattern Invoke() as fall-back, in these cases clickCount is ignored. If but is specified (for example "left", "right") then the native mouse Click function will be used to click the center of the element.
		;StringLower, but, but		
		if (but == "") {
			if (this.GetCurrentPropertyValue(UIA_Enum.UIA_IsInvokePatternAvailablePropertyId))
				this.GetCurrentPatternAs("Invoke").Invoke()
			else if (this.GetCurrentPropertyValue(UIA_Enum.UIA_IsTogglePatternAvailablePropertyId))
				this.GetCurrentPatternAs("Toggle").Toggle()
			else if (this.GetCurrentPropertyValue(UIA_Enum.UIA_IsExpandCollapsePatternAvailablePropertyId)) {
				if (this.GetCurrentPatternAs("ExpandCollapse").CurrentExpandCollapseState == 0)
					pattern.Expand()
				Else
					pattern.Collapse()
			} else if (this.GetCurrentPropertyValue(UIA_Enum.UIA_IsLegacyIAccessiblePatternAvailablePropertyId))
				this.GetCurrentPatternAs("LegacyIAccessible").DoDefaultAction()
		} else {
			pos := this.CurrentPos() ; or should GetClickablePos be used instead?
			Click, % pos.x+pos.w//2 " " pos.y+pos.h//2 " " but " " clickCount
		} 
			
	}
	CurrentPos(relativeTo="") { ; Returns an object containing the x, y coordinates and width and height: {x:x coordinate, y:y coordinate, w:width, h:height}. relativeTo can be client, window or screen, default is A_CoordModeMouse.
		relativeTo := (relativeTo == "") ? A_CoordModeMouse : relativeTo
		StringLower, relativeTo, relativeTo
		br := this.CurrentBoundingRectangle
		if (relativeTo == "screen")
			return {x:br.l, y:br.t, w:(br.r-br.l), h:(br.b-br.t)}
		else {
			hwnd := this.GetParentHwnd()
			;WinGetTitle, wTitle, ahk_id %hwnd% ; for debugging purposes
			;ToolTip, %wTitle%
			if ((relativeTo == "window") || (relativeTo == "relative")) {
				VarSetCapacity(RECT, 16)
				DllCall("user32\GetWindowRect", "Ptr", hwnd, "Ptr", &RECT)
				return {x:(br.l-NumGet(&RECT, 0, "Int")), y:(br.t-NumGet(&RECT, 4, "Int")), w:(br.r-br.l), h:(br.b-br.t)}
			} else if (relativeTo == "client") {
				VarSetCapacity(pt,8,0), NumPut(br.l,&pt,0,"int"), NumPut(br.t,&pt,4,"int")
				DllCall("ScreenToClient", "Ptr",hwnd, "Ptr",&pt)
				return {x:NumGet(pt,"int"), y:NumGet(pt,4,"int"), w:(br.r-br.l), h:(br.b-br.t)}
			}
		}			
	}
	GetChildren(scope=0x2) { ; By default get only direct children (UIA_TreeScope_Children)
		return this.FindAll(this.TrueCondition, scope)
	}
	TWGetChildren() { ; Get all child elements using TreeViewer
		arr := []
		if !IsObject(nextChild := this.TreeWalkerTrue.GetFirstChildElement(this))
			return 0
		arr.Push(nextChild)
		while IsObject(nextChild := this.TreeWalkerTrue.GetNextSiblingElement(nextChild))
			arr.Push(nextChild)
		return arr
	}
	TWRecursive(maxDepth=20, layer="") { ; This function might hang if the element has thousands of empty custom elements (e.g. complex webpage)
		StrReplace(layer, ".",, dotcount)
		if (dotcount > maxDepth)
			return ""
		if !(children := this.GetChildren())
			return ""
		returnStr := ""
		for k, v in children {
			returnStr .= layer . (layer == "" ? "" : ".") . k " " v.Dump() . "`n" . v.TWRecursive(maxDepth, layer (layer == "" ? "" : ".") k)
		}
		return returnStr
	}
	Dump() { ; Returns info about the element: ControlType, Name, Value, LocalizedControlType, AcceleratorKey. 
		return "Type: " this.CurrentControlType ((name := this.CurrentName) ? " Name: """ name """" : "") ((val := this.CurrentValue) ? " Value: """ val """": "") ((lct := this.CurrentLocalizedControlType) ? " LocalizedControlType: """ lct """" : "") ((ak := this.CurrentAcceleratorKey) ? " AcceleratorKey: """ ak """": "")
	}
	DumpAll(maxDepth=20) { ; Returns info (ControlType, Name etc) for all descendants of the element. maxDepth is the allowed depth of recursion, by default 20 layers. DO NOT call this on the root element!
		return this.TWRecursive(maxDepth)
	}
	/*
		FindFirst using search criteria. Scope by default is UIA_TreeScope_Descendants. If using Name as a criteria, matchMode follows SetTitleMatchMode scheme: 1=name must must start with the specified name; 2=can contain anywhere; 3=exact match; RegEx. The Name can't be empty.
		
		For a specific criteria the following syntax is used: PropertyId=matchvalue
		Example1: "Name=Username:" would use FindFirst with UIA_Enum.UIA_NamePropertyId looking for the string "Username:"
		Example2: "ControlType=Button would FindFirst using UIA_Enum.UIA_ControlTypePropertyId and UIA_ButtonControlTypeId. Alternatively "ControlType=50000" can be used (direct value for UIA_ButtonControlTypeId which is 50000)
		
		Criteria can be combined with AND, OR, &&, ||:
		Example3: "Name=Username: AND ControlType=Button" would FindFirst an element with the name property of "Username:" and control type of button.
		Parenthesis are not supported! Criteria are evaluated left to right, so "a AND b OR c" would be evaluated as "(a and b) or c".
		
		Negation can be specified with NOT:
		Example4: "NOT ControlType=Edit" would return the first element that is not an edit element
	*/
	FindFirstBy(expr, scope=0x4, matchMode=3, caseSensitive=True) { 
		pos := 1, match := "", createCondition := "", operator := "", bufName := []
		while (pos := RegexMatch(expr, "(.*?)=(.*?)( AND | OR | && | \|\| |$)", match, pos+StrLen(match))) {
			if !match
				break
			if (InStr(match1, "Name") && (matchMode != 2) && (matchMode != 3)) {
				bufName[1] := match1, bufName[2] := match2
				Continue
			} else {
				;MsgBox, % "Creating condition with: m1: """ match1 """ m2: """ match2 """ m3: """ match3 """ flags: " ((matchMode==2)?2:0)|!caseSensitive
				newCondition := (SubStr(match1, 1, 4) == "NOT ") ? this.__UIA.CreateNotCondition(this.__UIA.CreateCondition(SubStr(match1, 5), match2)) : this.__UIA.CreateCondition(match1, match2, ((matchMode==2 && match1=="Name")?2:0)|!caseSensitive)
			}
			fullCondition := (operator == " AND " || operator == " && ") ? this.__UIA.CreateAndCondition(fullCondition, newCondition) : (operator == " OR " || operator == " || ") ? this.__UIA.CreateOrCondition(fullCondition, newCondition) : newCondition
			operator := match3
		}
		if (bufName[1]) {
			notProp := !InStr(bufName[1], "NOT "), name := bufName[2]
			nameCondition := (matchMode==1)?this.__UIA.CreatePropertyConditionEx(UIA_Enum.UIA_NamePropertyId, name,, 2|!caseSensitive):this.__UIA.CreateNotCondition(this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_NamePropertyId, ""))
			fullCondition := IsObject(fullCondition) ? this.__UIA.CreateAndCondition(nameCondition, fullCondition) : nameCondition
			for k, v in this.FindAll(fullCondition, scope) {
				curName := v.CurrentName
				if notProp {
					if (((matchMode == 1) && (SubStr(curName, 1, StrLen(name)) == name)) || ((matchMode == "RegEx") && RegExMatch(curName, name)))
						return v
				} else {
					if (((matchMode == 1) && (SubStr(curName, 1, StrLen(name)) != name)) || ((matchMode == "RegEx") && !RegExMatch(curName, name)))
						return v
				}
			}
		} else {
			return this.FindFirst(fullCondition, scope)
		}
	}
	
	FindFirstByName(name, scope=0x4, matchMode=3, caseSensitive=True) { ; MatchMode has same convention as window TitleMatchMode: 1=needs to start with the specified name, 2=can contain anywhere, 3=exact match, RegEx=regex match
		if (matchMode == 3 || matchMode == 2) {
			;MsgBox, % "Find: " UIA_Enum.NamePropertyId " ControlType: " name
			nameCondition := this.__UIA.CreatePropertyConditionEx(UIA_Enum.UIA_NamePropertyId, name,, ((matchMode==3)?0:2)|!caseSensitive)
			return this.FindFirst(nameCondition, scope)
		}
		nameCondition := (matchMode==1)?this.__UIA.CreatePropertyConditionEx(UIA_Enum.UIA_NamePropertyId, name,, 2|!caseSensitive):this.__UIA.CreateNotCondition(this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_NamePropertyId, ""))
		for k, v in this.FindAll(nameCondition, scope) {
			curName := v.CurrentName
			if (((matchMode == 1) && (SubStr(curName, 1, StrLen(name)) == name)) || ((matchMode == "RegEx") && RegExMatch(curName, name)))
				return v		
		}
	}
	
	FindFirstByType(controlType, scope=0x4) {
		if controlType is not integer
			controlType := UIA_Enum.UIA_ControlTypeId(controlType)
		if !controlType
			throw Exception("Invalid control type specified", -1)
		controlCondition := this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_ControlTypePropertyId, controlType)
		return this.FindFirst(controlCondition, scope)
	}
	
	FindFirstByNameAndType(name, controlType, scope=0x4, matchMode=3, caseSensitive=True) {
		if controlType is not integer
			controlType := UIA_Enum.UIA_ControlTypeId(controlType)
		if !controlType
			throw Exception("Invalid control type specified", -1)
		controlCondition := this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_ControlTypePropertyId, controlType, 3)
		if (matchMode == 3 || matchMode == 2) {
			nameCondition := this.__UIA.CreatePropertyConditionEx(UIA_Enum.UIA_NamePropertyId, name,, ((matchMode==3)?0:2)|!caseSensitive)
			AndCondition := this.__UIA.CreateAndCondition(nameCondition, controlCondition)
			return this.FindFirst(AndCondition, scope)
		}
		nameCondition := (matchMode==1)?this.__UIA.CreatePropertyConditionEx(UIA_Enum.UIA_NamePropertyId, name,, 2|!caseSensitive):this.__UIA.CreateNotCondition(this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_NamePropertyId, ""))
		AndCondition := this.__UIA.CreateAndCondition(nameCondition, controlCondition)
		for k, v in this.FindAll(AndCondition, scope) {
			curName := v.CurrentName
			if (((matchMode == 1) && (SubStr(curName, 1, StrLen(name)) == name)) || ((matchMode == "RegEx") && RegExMatch(curName, name)))
				return v		
		}
	}

	FindAllBy(expr, scope=0x4, matchMode=3, caseSensitive=True) { 
		pos := 1, match := "", createCondition := "", operator := "", bufName := []
		while (pos := RegexMatch(expr, "(.*?)=(.*?)( AND | OR | && | \|\| |$)", match, pos+StrLen(match))) {
			if !match
				break
			if (InStr(match1, "Name") && (matchMode != 2) && (matchMode != 3)) {
				bufName[1] := match1, bufName[2] := match2
				Continue
			} else {
				;MsgBox, % "Creating condition with: m1: """ match1 """ m2: """ match2 """ m3: """ match3 """ flags: " ((matchMode==2)?2:0)|!caseSensitive
				newCondition := (SubStr(match1, 1, 4) == "NOT ") ? this.__UIA.CreateNotCondition(this.__UIA.CreateCondition(SubStr(match1, 5), match2)) : this.__UIA.CreateCondition(match1, match2, ((matchMode==2 && match1=="Name")?2:0)|!caseSensitive)
			}
			fullCondition := (operator == " AND " || operator == " && ") ? this.__UIA.CreateAndCondition(fullCondition, newCondition) : (operator == " OR " || operator == " || ") ? this.__UIA.CreateOrCondition(fullCondition, newCondition) : newCondition
			operator := match3
		}
		if (bufName[1]) {
			notProp := !InStr(bufName[1], "NOT "), name := bufName[2], returnArr := []
			nameCondition := (matchMode==1)?this.__UIA.CreatePropertyConditionEx(UIA_Enum.UIA_NamePropertyId, name,, 2|!caseSensitive):this.__UIA.CreateNotCondition(this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_NamePropertyId, ""))
			fullCondition := IsObject(fullCondition) ? this.__UIA.CreateAndCondition(nameCondition, fullCondition) : nameCondition
			for k, v in this.FindAll(fullCondition, scope) {
				curName := v.CurrentName
				if notProp {
					if (((matchMode == 1) && (SubStr(curName, 1, StrLen(name)) == name)) || ((matchMode == "RegEx") && RegExMatch(curName, name)))
						returnArr.Push(v)
				} else {
					if (((matchMode == 1) && (SubStr(curName, 1, StrLen(name)) != name)) || ((matchMode == "RegEx") && !RegExMatch(curName, name)))
						returnArr.Push(v)
				}
			}
			return returnArr
		} else {
			return this.FindAll(fullCondition, scope)
		}
	}

	FindAllByName(name, scope=0x4, matchMode=3, caseSensitive=True) {
		if (matchMode == 3 || matchMode == 2) {
			nameCondition := this.__UIA.CreatePropertyConditionEx(UIA_Enum.UIA_NamePropertyId, name,, ((matchMode==3)?0:2)|!caseSensitive)
			return this.FindAll(nameCondition, scope)
		}
		nameCondition := (matchMode==1)?this.__UIA.CreatePropertyConditionEx(UIA_Enum.UIA_NamePropertyId, name,, 2|!caseSensitive):this.__UIA.CreateNotCondition(this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_NamePropertyId, ""))
		retList := []
		for k, v in this.FindAll(nameCondition, scope) {
			curName := v.CurrentName
			if (((matchMode == 1) && (SubStr(curName, 1, StrLen(name)) == name))|| ((matchMode == "RegEx") && RegExMatch(curName, name)))
				retList.Push(v)		
		}
		return retList
	}

	FindAllByType(controlType, scope=0x4) {
		if controlType is not integer
			controlType := UIA_Enum.UIA_ControlTypeId(controlType)
		if !controlType
			throw Exception("Invalid control type specified", -1)
		controlCondition := this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_ControlTypePropertyId, controlType)
		return this.FindAll(controlCondition, scope)
	}

	FindAllByNameAndType(name, controlType, scope=0x4, matchMode=3) {
		if controlType is not integer
			controlType := UIA_Enum.UIA_ControlTypeId(controlType)
		if !controlType
			throw Exception("Invalid control type specified", -1)
		controlCondition := this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_ControlTypePropertyId, controlType)
		if (matchMode == 3 || matchMode == 2) {
			nameCondition := this.__UIA.CreatePropertyConditionEx(UIA_Enum.UIA_NamePropertyId, name, ((matchMode==3)?0:2)|!caseSensitive)
			AndCondition := this.__UIA.CreateAndCondition(nameCondition, controlCondition)
			return this.FindAll(AndCondition, scope)
		}
		nameCondition := (matchMode==1)?this.__UIA.CreatePropertyConditionEx(UIA_Enum.UIA_NamePropertyId, name, , 2|!caseSensitive):this.__UIA.CreateNotCondition(this.__UIA.CreatePropertyCondition(UIA_Enum.UIA_NamePropertyId, ""))
		AndCondition := this.__UIA.CreateAndCondition(nameCondition, controlCondition)
		returnArr := []
		for k, v in this.FindAll(AndCondition, scope) {
			curName := v.CurrentName
			if (((matchMode == 1) && (SubStr(curName, 1, StrLen(name)) == name)) || ((matchMode == "RegEx") && RegExMatch(curName, name)))
				returnArr.Push(v)	
		}
		return returnArr
	}

	FindByPath(searchPath="") { ; Gets an element by the "path" that is displayed in the UIA_Element.DumpAll() result. This is like the Acc path, but for UIA (they are not compatible).
		el := this
		Loop, Parse, searchPath, .
		{
			children := el.GetChildren()
			if !IsObject(el := children[A_LoopField])
				return
		}
		return el
	}
	
	WaitElementExist(expr, scope=0x4, matchMode=3, caseSensitive=True, timeOut=10000) { ; Calls UIA_Element.FindFirstBy until the element is found and then returns it, with a timeOut of 10000ms (10 seconds)
		startTime := A_TickCount
		while (!IsObject(el := this.FindFirstBy(expr, scope, matchMode, caseSensitive)) && ((timeOut < 1) ? 1 : (A_tickCount - startTime < timeOut)))
			Sleep, 100
		return el
	}
	WaitElementExistByName(name, scope=0x4, matchMode=3, caseSensitive=True, timeOut=10000) {
		startTime := A_TickCount
		while (!IsObject(el := this.FindFirstByName(name, scope, matchMode, caseSensitive)) && ((timeOut < 1) ? 1 : (A_tickCount - startTime < timeOut)))
			Sleep, 100
		return el
	}
	WaitElementExistByType(controlType, scope=0x4, timeOut=10000) {
		startTime := A_TickCount
		while (!IsObject(el := this.FindFirstByType(controlType, scope)) && ((timeOut < 1) ? 1 : (A_tickCount - startTime < timeOut)))
			Sleep, 100
		return el
	}
	WaitElementExistByNameAndType(name, controlType, scope=0x4, matchMode=3, caseSensitive=True, timeOut=10000) {
		startTime := A_TickCount
		while (!IsObject(el := this.FindFirstByNameAndType(name, controlType, scope, matchMode, caseSensitive)) && ((timeOut < 1) ? 1 : (A_tickCount - startTime < timeOut))) {
			Sleep, 100
		}
		return el
	}
}

class UIA_Element2 extends UIA_Element {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671425(v=vs.85).aspx
	static __IID := "{6749C683-F70D-4487-A698-5F79D55290D6}"
		,  __properties := UIA_Element.__properties "`r`nCurrentOptimizeForVisualContent,85,BOOL`r`nCachedOptimizeForVisualContent,86,BOOL`r`nCurrentLiveSetting,87,int`r`nCachedLiveSetting,88,int`r`nCurrentFlowsFrom,89,UIA_ElementArray`r`nCachedFlowsFrom,90,UIA_ElementArray`r`n"
}
class UIA_Element3 extends UIA_Element2 {
	static __IID := "{8471DF34-AEE0-4A01-A7DE-7DB9AF12C296}"
		,  __properties := UIA_Element2.__properties "`r`nCurrentIsPeripheral,92,int`r`nCachedIsPeripheral,93,int"
	ShowContextMenu() {
		return UIA_Hr(DllCall(this.__Vt(91), "ptr",this.__Value))
	}
}
class UIA_Element4 extends UIA_Element3 {
	static __IID := "{3B6E233C-52FB-4063-A4C9-77C075C2A06B}"
		,  __properties := UIA_Element3.__properties "`r`nCurrentPositionInSet,94,int`r`nCurrentSizeOfSet,95,int`r`nCurrentLevel,96,int`r`nCurrentAnnotationTypes,97,VARIANT`r`nCurrentAnnotationObjects,98,IUIAutomationElementArray`r`nCachedPositionInSet,99,int`r`nCachedSizeOfSet,100,int`r`nCachedLevel,101,int`r`nCachedAnnotationTypes,102,VARIANT`r`nCachedAnnotationObjects,103,IUIAutomationElementArray"
}
class UIA_Element5 extends UIA_Element4 {
	static __IID := "{98141C1D-0D0E-4175-BBE2-6BFF455842A7}"
		,  __properties := UIA_Element4.__properties "`r`nCurrentLandmarkType,104,int`r`nCurrentLocalizedLandmarkType,105,BSTR`r`nCachedLandmarkType,106,int`r`nCachedLocalizedLandmarkType,107,BSTR"
}
class UIA_Element6 extends UIA_Element5 {
	static __IID := "{4780D450-8BCA-4977-AFA5-A4A517F555E3}"
		,  __properties := UIA_Element5.__properties "`r`nCurrentFullDescription,108,BSTR`r`nCachedFullDescription,109,BSTR"
}
class UIA_Element7 extends UIA_Element6 {
	static __IID := "{204E8572-CFC3-4C11-B0C8-7DA7420750B7}"
		,  __properties := UIA_Element6.__properties
	FindFirstWithOptions(scope, c, traversalOptions=0, root=0) {
		return UIA_Hr(DllCall(this.__Vt(110), "ptr",this.__Value, "uint",scope, "ptr",(c=""?this.TrueCondition:c).__Value, "int", traversalOptions, "ptr", root.__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	FindAllWithOptions(scope, c, traversalOptions=0, root=0) {
		return UIA_Hr(DllCall(this.__Vt(111), "ptr",this.__Value, "uint",scope, "ptr",(c=""?this.TrueCondition:c).__Value, "int", traversalOptions, "ptr", root.__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	FindFirstWithOptionsBuildCache(scope, c, cacheRequest, traversalOptions=0, root=0) {
		return UIA_Hr(DllCall(this.__Vt(112), "ptr",this.__Value, "uint",scope, "ptr",(c=""?this.TrueCondition:c).__Value, "ptr", cacheRequest.__Value, "int", traversalOptions, "ptr", root.__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	FindAllWithOptionsBuildCache(scope, c, cacheRequest, traversalOptions=0, root=0) {
		return UIA_Hr(DllCall(this.__Vt(113), "ptr",this.__Value, "uint",scope, "ptr",(c=""?this.TrueCondition:c).__Value, "ptr", cacheRequest.__Value, "int", traversalOptions, "ptr", root.__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	GetCurrentMetadataValue(targetId, metadataId) {
		return UIA_Hr(DllCall(this.__Vt(114), "ptr",this.__Value, "int",targetId, "int", metadataId, "ptr*", UIA_Variant(out)))? UIA_VariantData(out):
	}
}

class UIA_ElementArray extends UIA_Base {
	static __IID := "{14314595-b4bc-4055-95f2-58f2e42c9855}"
		,  __properties := "Length,3,int"
	
	GetElement(i) {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "int",i, "ptr*",out))? UIA_Element(out):
	}
}

class UIA_TreeWalker extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671470(v=vs.85).aspx
	static __IID := "{4042c624-389c-4afc-a630-9df854a541fc}"
		,  __properties := "Condition,15,IUIAutomationCondition"
	
	GetParentElement(e) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))? UIA_Element(out):
	}
	GetFirstChildElement(e) {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	GetLastChildElement(e) {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	GetNextSiblingElement(e) {
		return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	GetPreviousSiblingElement(e) {
		return UIA_Hr(DllCall(this.__Vt(7), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	NormalizeElement(e) {
		return UIA_Hr(DllCall(this.__Vt(8), "ptr",this.__Value, "ptr",e.__Value, "ptr*",out))&&out? UIA_Element(out):
	}
	GetParentElementBuildCache(e, cacheRequest) {
		return UIA_Hr(DllCall(this.__Vt(9), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}
	GetFirstChildElementBuildCache(e, cacheRequest) {
		return UIA_Hr(DllCall(this.__Vt(10), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}
	GetLastChildElementBuildCache(e, cacheRequest) {
		return UIA_Hr(DllCall(this.__Vt(11), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}
	GetNextSiblingElementBuildCache(e, cacheRequest) {
		return UIA_Hr(DllCall(this.__Vt(12), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}
	GetPreviousSiblingElementBuildCache(e, cacheRequest) {
		return UIA_Hr(DllCall(this.__Vt(13), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}
	NormalizeElementBuildCache(e, cacheRequest) {
		return UIA_Hr(DllCall(this.__Vt(14), "ptr",this.__Value, "ptr",e.__Value, "ptr",cacheRequest.__Value, "ptr*",out))? UIA_Element(out):
	}
}

class UIA_Condition extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671420(v=vs.85).aspx
	static __IID := "{352ffba8-0973-437c-a61f-f64cafd81df9}"
		,  __properties := ""
}

class UIA_PropertyCondition extends UIA_Condition {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696121(v=vs.85).aspx
	static __IID := "{99ebf2cb-5578-4267-9ad4-afd6ea77e94b}"
		,  __properties := "PropertyId,3,PROPERTYID`r`nPropertyValue,4,VARIANT`r`nPropertyConditionFlags,5,PropertyConditionFlags"
}

class UIA_AndCondition extends UIA_Condition {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671407(v=vs.85).aspx
	static __IID := "{a7d0af36-b912-45fe-9855-091ddc174aec}"
		,  __properties := "ChildCount,3,int"
	
	;~ GetChildrenAsNativeArray	4	IUIAutomationCondition ***childArray
	GetChildren() { ; Returns a native AHK array containing all the conditions (already subtyped to AndCondition, OrCondition etc)
		ret := UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr*",out)), arr := []
		if (out && (safeArray := ComObj(0x2003,out,1))) {
			for k in safeArray {
				obj := ComObject(9, k, 1), ObjAddRef(k)
				for _, n in ["Property", "Bool", "And", "Or", "Not"] {
					if ComObjQuery(obj, UIA_%n%Condition.__IID) {
						arr.Push(new UIA_%n%Condition(k))
						break
					}
				}
				ObjRelease(k)
			}
			return arr
		}
		return
	}
}
class UIA_OrCondition extends UIA_Condition {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696108(v=vs.85).aspx
	static __IID := "{8753f032-3db1-47b5-a1fc-6e34a266c712}"
		,  __properties := "ChildCount,3,int"
	
	;~ GetChildrenAsNativeArray	4	IUIAutomationCondition ***childArray
	GetChildren() {
		ret := UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr*",out)), arr := []
		if (out && (safeArray := ComObject(0x2003,out,1))) {
			for k in safeArray {
				obj := ComObject(9, k, 1), ObjAddRef(k)
				for _, n in ["Property", "Bool", "And", "Or", "Not"] {
					if ComObjQuery(obj, UIA_%n%Condition.__IID) {
						arr.Push(new UIA_%n%Condition(k))
						break
					}
				}
				ObjRelease(k)
			}
			return arr
		}
		return
	}
}
class UIA_BoolCondition extends UIA_Condition {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671411(v=vs.85).aspx
	static __IID := "{1B4E1F2E-75EB-4D0B-8952-5A69988E2307}"
		,  __properties := "BooleanValue,3,boolVal"
}

class UIA_NotCondition extends UIA_Condition {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696106(v=vs.85).aspx
	static __IID := "{f528b657-847b-498c-8896-d52b565407a1}"

	GetChild() { ; Type of the received condition can be determined with out.__Class
		ret := UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr*",out)), obj := ComObject(9, out, 1)
		for k, v in ["Bool", "Property", "And", "Or", "Not"] {
			if ComObjQuery(obj, UIA_%v%Condition.__IID)
				return ret?new UIA_%v%Condition(out):
		}
		return UIA_Hr(0x80004005)
	}
}

class UIA_IUnknown extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ms680509(v=vs.85).aspx
	static __IID := "{00000000-0000-0000-C000-000000000046}"
}

class UIA_CacheRequest extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671413(v=vs.85).aspx
	static __IID := "{b32a92b5-bc25-4078-9c08-d7ee95c48e03}"
		,  __properties := "CurrentTreeScope,6,int`r`nCurrentTreeFilter,8,int`r`nCurrentAutomationElementMode,10,int"

	Clone() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr",out)) ? new UIA_CacheRequest(out):
	}	
	AddProperty(property) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr",property))
	}
	AddPattern(pattern) {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value,"ptr",pattern))
	}
	SetTreeScope(scope) {
		return UIA_Hr(DllCall(this.__Vt(7), "ptr",this.__Value, "ptr",scope))
	}
	SetTreeFilter(condition) {
		return UIA_Hr(DllCall(this.__Vt(9), "ptr",this.__Value, "ptr",condition.__Value))
	}
	SetAutomationElementMode(mode) {
		return UIA_Hr(DllCall(this.__Vt(11), "ptr",this.__Value, "ptr",mode))
	}
}

class _UIA_EventHandler extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696044(v=vs.85).aspx
	static __IID := "{146c3c17-f12e-4e22-8c27-f894b9b79c69}"

	HandleAutomationEvent(sender, eventId) {
		param1 := this, this := Object(A_EventInfo), funcName := this.__Version
		%funcName%(UIA_Element(sender), eventId)
		return param1
	}
}
class _UIA_FocusChangedEventHandler extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696051(v=vs.85).aspx
	static __IID := "{c270f6b5-5c69-4290-9745-7a7f97169468}"
	HandleFocusChangedEvent(sender) {
		param1 := this, this := Object(A_EventInfo), funcName := this.__Version
		%funcName%(UIA_Element(sender))
		return param1
	}
}
class _UIA_PropertyChangedEventHandler extends UIA_Base { ; UNTESTED
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696119(v=vs.85).aspx
	static __IID := "{40cd37d4-c756-4b0c-8c6f-bddfeeb13b50}"
	HandlePropertyChangedEvent(sender, propertyId, newValue) {
		param1 := this, this := Object(A_EventInfo), funcName := this.__Version
		%funcName%(UIA_Element(sender), eventId, UIA_VariantData(newValue))
		return param1
	}
}
class _UIA_StructureChangedEventHandler extends UIA_Base { ; UNTESTED
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696197(v=vs.85).aspx
	static __IID := "{e81d1b4e-11c5-42f8-9754-e7036c79f054}"
	HandleStructureChangedEvent(sender, changeType, runtimeId) {
		param1 := this, this := Object(A_EventInfo), funcName := this.__Version
		%funcName%(UIA_Element(sender), changeType, UIA_SafeArrayToAHKArray(ComObj(0x2003,runtimeId,1)))
		return param1
	}
}
class _UIA_TextEditTextChangedEventHandler { ; UNTESTED
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/dn302202(v=vs.85).aspx
	static __IID := "{92FAA680-E704-4156-931A-E32D5BB38F3F}"
	HandleTextEditTextChangedEvent(sender, changeType, eventStrings) {
		param1 := this, this := Object(A_EventInfo), funcName := this.__Version
		%funcName%(UIA_Element(sender), changeType, UIA_SafeArrayToAHKArray(ComObj(0x2008,eventStrings,1)))
		return param1
	}
	;~ HandleTextEditTextChangedEvent	3
}

class _UIA_ChangesEventHandler { ; UNTESTED
	static __IID := "{58EDCA55-2C3E-4980-B1B9-56C17F27A2A0}"
	HandleChangesEvent(sender, uiaChanges, changesCount) {
		param1 := this, this := Object(A_EventInfo), funcName := this.__Version, changes := {}
		changes.uiaId := NumGet(uiaChanges, 0), changes.payload := UIA_VariantData(uiaChanges, 8), changes.extraInfo := UIA_VariantData(uiaChanges,16+2*A_PtrSize)
		
		%funcName%(UIA_Element(sender), changes, changesCount)
		return param1
	}
}

class _UIA_NotificationEventHandler {
	static __IID := "{C7CB2637-E6C2-4D0C-85DE-4948C02175C7}"
	HandleNotificationEvent(sender, notificationKind, notificationProcessing, displayString, activityId) {
		param1 := this, this := Object(A_EventInfo), funcName := this.__Version
		%funcName%(UIA_Element(sender), notificationKind, notificationProcessing, StrGet(displayString), StrGet(activityId))
		return param1
	}
}


;~ 		UIA_Patterns - http://msdn.microsoft.com/en-us/library/windows/desktop/ee684023
class UIA_DockPattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee671421
	static	__IID := "{fde5ef97-1464-48f6-90bf-43d0948e86ec}"
		,	__PatternID := 10011
		,	__Properties := "CurrentDockPosition,4,int`r`nCachedDockPosition,5,int"

	SetDockPosition(Pos) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "uint",pos))
	}
	/*	DockPosition_Top	= 0,
		DockPosition_Left	= 1,
		DockPosition_Bottom	= 2,
		DockPosition_Right	= 3,
		DockPosition_Fill	= 4,
		DockPosition_None	= 5
	*/
}
class UIA_ExpandCollapsePattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696046
	static	__IID := "{619be086-1f4e-4ee4-bafa-210128738730}"
		,	__PatternID := 10005
		,	__Properties := "CachedExpandCollapseState,6,int`r`nCurrentExpandCollapseState,5,int"
	
	Expand() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
	}
	Collapse() {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value))
	}	
	/*	ExpandCollapseState_Collapsed	= 0,
		ExpandCollapseState_Expanded	= 1,
		ExpandCollapseState_PartiallyExpanded	= 2,
		ExpandCollapseState_LeafNode	= 3
	*/
}
class UIA_GridItemPattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696053
	static	__IID := "{78f8ef57-66c3-4e09-bd7c-e79b2004894d}"
		,	__PatternID := 10007
		,	__Properties := "CurrentContainingGrid,3,IUIAutomationElement`r`nCurrentRow,4,int`r`nCurrentColumn,5,int`r`nCurrentRowSpan,6,int`r`nCurrentColumnSpan,7,int`r`nCachedContainingGrid,8,IUIAutomationElement`r`nCachedRow,9,int`r`nCachedColumn,10,int`r`nCachedRowSpan,11,int`r`nCachedColumnSpan,12,int"
}
class UIA_GridPattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696064
	static	__IID := "{414c3cdc-856b-4f5b-8538-3131c6302550}"
		,	__PatternID := 10006
		,	__Properties := "CurrentRowCount,4,int`r`nCurrentColumnCount,5,int`r`nCachedRowCount,6,int`r`nCachedColumnCount,7,int"

	GetItem(row,column) { ; Hr!=0 if no result, or blank output?
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "uint",row, "uint",column, "ptr*",out))? UIA_Element(out):
	}
}
class UIA_InvokePattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696070
	static	__IID := "{fb377fbe-8ea6-46d5-9c73-6499642d3059}"
		,	__PatternID := 10000
	
	Invoke() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
	}
}
class UIA_ItemContainerPattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696072
	static	__IID := "{c690fdb2-27a8-423c-812d-429773c9084e}"
		,	__PatternID := 10019

	FindItemByProperty(startAfter, propertyId, ByRef value, type=8) {	; Hr!=0 if no result, or blank output?
		if (type!="Variant")
			UIA_Variant(value,type,value)
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr",startAfter.__Value, "int",propertyId, "ptr",&value, "ptr*",out))? UIA_Element(out):
	}
}
class UIA_LegacyIAccessiblePattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696074
	static	__IID := "{828055ad-355b-4435-86d5-3b51c14a9b1b}"
		,	__PatternID := 10018
		,	__Properties := "CurrentChildId,6,int`r`nCurrentName,7,BSTR`r`nCurrentValue,8,BSTR`r`nCurrentDescription,9,BSTR`r`nCurrentRole,10,DWORD`r`nCurrentState,11,DWORD`r`nCurrentHelp,12,BSTR`r`nCurrentKeyboardShortcut,13,BSTR`r`nCurrentDefaultAction,15,BSTR`r`nCachedChildId,16,int`r`nCachedName,17,BSTR`r`nCachedValue,18,BSTR`r`nCachedDescription,19,BSTR`r`nCachedRole,20,DWORD`r`nCachedState,21,DWORD`r`nCachedHelp,22,BSTR`r`nCachedKeyboardShortcut,23,BSTR`r`nCachedDefaultAction,25,BSTR"

	Select(flags=3) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "int",flags))
	}
	DoDefaultAction() {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value))
	}
	SetValue(value) {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr",&value))
	}
	GetCurrentSelection() { ; UNTESTED
		return UIA_Hr(DllCall(this.__Vt(14), "ptr",this.__Value, "ptr*",out))&&out? UIA_ElementArray(out):
	}
	;~ GetCachedSelection	24	IUIAutomationElementArray
	GetIAccessible() {
	/*	This method returns NULL if the underlying implementation of the UI Automation element is not a native 
	Microsoft Active Accessibility server; that is, if a client attempts to retrieve the IAccessible interface 
	for an element originally supported by a proxy object from OLEACC.dll, or by the UIA-to-MSAA Bridge.
	*/
		return UIA_Hr(DllCall(this.__Vt(26), "ptr",this.__Value, "ptr*",pacc))&&pacc? ComObj(9,pacc,1):
	}
}
class UIA_MultipleViewPattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696099
	static	__IID := "{8d253c91-1dc5-4bb5-b18f-ade16fa495e8}"
		,	__PatternID := 10008
		,	__Properties := "CurrentCurrentView,5,int`r`nCachedCurrentView,7,int"

	GetViewName(view) { ; need to release BSTR?
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "int",view, "ptr*",name))? StrGet(name):
	}
	SetCurrentView(view) {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "int",view))
	}
	GetCurrentSupportedViews() {
		return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "ptr*",out))? UIA_SafeArrayToAHKArray(ComObj(0x2003,out,1)):
	}
	GetCachedSupportedViews() {
		return UIA_Hr(DllCall(this.__Vt(8), "ptr",this.__Value, "ptr*",out))? UIA_SafeArrayToAHKArray(ComObj(0x2003,out,1)):
	}
}
class UIA_RangeValuePattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696147
	static	__IID := "{59213f4f-7346-49e5-b120-80555987a148}"
		,	__PatternID := 10003
		,	__Properties := "CurrentValue,4,double`r`nCurrentIsReadOnly,5,BOOL`r`nCurrentMaximum,6,double`r`nCurrentMinimum,7,double`r`nCurrentLargeChange,8,double`r`nCurrentSmallChange,9,double`r`nCachedValue,10,double`r`nCachedIsReadOnly,11,BOOL`r`nCachedMaximum,12,double`r`nCachedMinimum,13,double`r`nCachedLargeChange,14,double`r`nCachedSmallChange,15,double"

	SetValue(val) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "double",val))
	}
}
class UIA_ScrollItemPattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696165
	static	__IID := "{b488300f-d015-4f19-9c29-bb595e3645ef}"
		,	__PatternID := 10017

	ScrollIntoView() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
	}
}
class UIA_ScrollPattern extends UIA_Base {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696167
	static	__IID := "{88f4d42a-e881-459d-a77c-73bbbb7e02dc}"
		,	__PatternID := 10004
		,	__Properties := "CurrentHorizontalScrollPercent,5,double`r`nCurrentVerticalScrollPercent,6,double`r`nCurrentHorizontalViewSize,7,double`r`nCurrentHorizontallyScrollable,9,BOOL`r`nCurrentVerticallyScrollable,10,BOOL`r`nCachedHorizontalScrollPercent,11,double`r`nCachedVerticalScrollPercent,12,double`r`nCachedHorizontalViewSize,13,double`r`nCachedVerticalViewSize,14,double`r`nCachedHorizontallyScrollable,15,BOOL`r`nCachedVerticallyScrollable,16,BOOL"
		
	Scroll(horizontal=-1, vertical=-1) { ; Default is NoScroll
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "uint",horizontal, "uint",vertical))
	}
	SetScrollPercent(horizontal=-1, vertical=-1) { ; Default is NoScroll
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "double",horizontal, "double",vertical))
	}
	/*	UIA_ScrollPatternNoScroll	=	-1
		ScrollAmount_LargeDecrement	= 0,
		ScrollAmount_SmallDecrement	= 1,
		ScrollAmount_NoAmount	= 2,
		ScrollAmount_LargeIncrement	= 3,
		ScrollAmount_SmallIncrement	= 4
	*/
}
class UIA_SelectionItemPattern extends UIA_Base { ; UNTESTED
	static	__IID := "{A8EFA66A-0FDA-421A-9194-38021F3578EA}"
		,	__PatternID := 10010
		,	__Properties := "CurrentIsSelected,6,int`r`nCurrentSelectionContainer,7,IUIAutomationElement`r`nCachedIsSelected,8,int`r`nCachedSelectionContainer,9,IUIAutomationElement"
	Select() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
	}
	AddToSelection() {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value))
	}
	RemoveFromSelection() {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value))
	}
}

class UIA_SelectionPattern extends UIA_Base {
	static	__IID := "{5ED5202E-B2AC-47A6-B638-4B0BF140D78E}"
		,	__PatternID := 10001
		,	__Properties := "CurrentCanSelectMultiple,4,int`r`nCurrentIsSelectionRequired,5,int`r`nCachedCanSelectMultiple,7,int`r`nCachedIsSelectionRequired,8,int"
	GetCurrentSelection() { ; returns an array of selected elements
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
	GetCachedSelection() {
		return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
}

class UIA_SelectionPattern2 extends UIA_SelectionPattern {
	static	__IID := "{0532bfae-c011-4e32-a343-6d642d798555"
		,	__PatternID := 10034
		,	__Properties := "CurrentFirstSelectedItem,9,IUIAutomationElement`r`nCurrentLastSelectedItem,10,IUIAutomationElement`r`nCurrentCurrentSelectedItem,11,IUIAutomationElement`r`nCurrentItemCount,12,int`r`nCachedFirstSelectedItem,13,IUIAutomationElement`r`nCachedLastSelectedItem,14,IUIAutomationElement`r`nCachedCurrentSelectedItem,15,IUIAutomationElement`r`nCachedItemCount,16,int"
	
}

class UIA_SpreadsheetItemPattern extends UIA_Base { ; UNTESTED
	static	__IID := "{7D4FB86C-8D34-40E1-8E83-62C15204E335}"
		,	__PatternID := 10027
		,	__Properties := "CurrentFormula,3,BSTR`r`nCachedFormula,6,BSTR"
	GetCurrentAnnotationObjects() {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
	GetCurrentAnnotationTypes() {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr", UIA_Variant(out)))?UIA_VariantData(out):
	}
	GetCachedAnnotationObjects() {
		return UIA_Hr(DllCall(this.__Vt(7), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
	GetCachedAnnotationTypes() {
		return UIA_Hr(DllCall(this.__Vt(8), "ptr",this.__Value, "ptr", UIA_Variant(out)))?UIA_VariantData(out):
	}
}

class UIA_SpreadsheetPattern extends UIA_Base { ; UNTESTED
	static	__IID := "{7517A7C8-FAAE-4DE9-9F08-29B91E8595C1}"
		,	__PatternID := 10026
		,	__Properties := ""
	GetItemByName(name) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr", &name, "ptr*", out))? UIA_Element(out):
	}
}

class UIA_StylesPattern extends UIA_Base { ; UNTESTED
	static	__IID := "{85B5F0A2-BD79-484A-AD2B-388C9838D5FB}"
		,	__PatternID := 10025
		,	__Properties := "CurrentStyleId,3,int`r`nCurrentStyleName,4,int`r`nCurrentFillColor,5,int`r`nCurrentFillPatternStyle,6,int`r`nCurrentShape,7,int`r`nCurrentFillPatternColor,8,int`r`nCurrentExtendedProperties,9,int`r`nCachedStyleId,11,int`r`nCachedStyleName,12,int`r`nCachedFillColor,13,int`r`nCachedFillPatternStyle,14,int`r`nCachedShape,15,int`r`nCachedFillPatternColor,16,int`r`nCachedExtendedProperties,17,int"
	GetCurrentExtendedPropertiesAsArray(byref propertyCount) {
		return UIA_Hr(DllCall(this.__Vt(10), "ptr",this.__Value, "ptr*", propertyArray, "int*", propertyCount))?UIA_SafeArrayToAHKArray(ComObj(0x2003,propertyArray,1)):
	}
	GetCachedExtendedPropertiesAsArray(byref propertyCount) {
		return UIA_Hr(DllCall(this.__Vt(18), "ptr",this.__Value, "ptr*", propertyArray, "int*", propertyCount))?UIA_SafeArrayToAHKArray(ComObj(0x2003,propertyArray,1)):
	}
}

class UIA_SynchronizedInputPattern extends UIA_Base { ; UNTESTED
	static	__IID := "{2233BE0B-AFB7-448B-9FDA-3B378AA5EAE1}"
		,	__PatternID := 10021
		,	__Properties := ""
	StartListening(inputType) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "int",inputType))
	}
	Cancel() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
	}
}
class UIA_TableItemPattern extends UIA_Base { ; UNTESTED
	static	__IID := "{0B964EB3-EF2E-4464-9C79-61D61737A27E}"
		,	__PatternID := 10013
		,	__Properties := ""
	GetCurrentRowHeaderItems() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
	GetCurrentColumnHeaderItems() {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
	GetCachedRowHeaderItems() {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
	GetCachedColumnHeaderItems() {
		return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
}
class UIA_TablePattern extends UIA_Base { ; UNTESTED
	static	__IID := "{620E691C-EA96-4710-A850-754B24CE2417}"
		,	__PatternID := 10012
		,	__Properties := "CurrentRowOrColumnMajor,5,int`r`nCachedRowOrColumnMajor,5,int"
	GetCurrentRowHeaders() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
	GetCurrentColumnHeaders() {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
	GetCachedRowHeaders() {
		return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
	GetCachedColumnHeaders() {
		return UIA_Hr(DllCall(this.__Vt(7), "ptr",this.__Value, "ptr*", out))&&out?UIA_ElementArray(out):
	}
}

class UIA_TextPattern extends UIA_Base {
	static	__IID := "{32EBA289-3583-42C9-9C59-3B6D9A1E9B6A}"
		,	__PatternID := 10014
		,	__Properties := "DocumentRange,7,IUIAutomationTextRange`r`nSupportedTextSelection,8,int"
	RangeFromPoint(x, y) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "int64",x&0xFFFFFFFF|y<<32, "ptr*",out))?UIA_TextRange(out):
	}
	RangeFromChild(child) {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "ptr",child.__Value, "ptr*",out))?UIA_TextRange(out):
	}
	GetSelection() {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "ptr*",out))?UIA_TextRangeArray(out):
	}
	GetVisibleRanges() {
		return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value, "ptr*",out))?UIA_TextRangeArray(out):
	}
}

class UIA_TextPattern2 extends UIA_TextPattern {
	static	__IID := "{506A921A-FCC9-409F-B23B-37EB74106872}"
		,	__PatternID := 10024
		,	__Properties := UIA_TextPattern.__Properties
	RangeFromAnnotation(annotation) {
		return UIA_Hr(DllCall(this.__Vt(9), "ptr",this.__Value, "ptr", annotation.__Value, "ptr*",out))?UIA_TextRange(out):
	}
	GetCaretRange(ByRef isActive="") {
		return UIA_Hr(DllCall(this.__Vt(10), "ptr",this.__Value, "ptr*", isActive, "ptr*",out))?UIA_TextRange(out):
	}
}

class UIA_TextEditPattern extends UIA_TextPattern { ; UNTESTED
	static	__IID := "{17E21576-996C-4870-99D9-BFF323380C06}"
		,	__PatternID := 10032
		,	__Properties := ""
	GetActiveComposition() {
		return UIA_Hr(DllCall(this.__Vt(9), "ptr",this.__Value, "ptr*", out))?UIA_TextRange(out):
	}
	GetConversionTarget() {
		return UIA_Hr(DllCall(this.__Vt(10), "ptr",this.__Value, "ptr*", out))?UIA_TextRange(out):
	}
}

class UIA_TextChildPattern extends UIA_Base { ; UNTESTED
	static	__IID := "{6552B038-AE05-40C8-ABFD-AA08352AAB86}"
		,	__PatternID := 10029
		,	__Properties := "TextContainer,3,IUIAutomationElement`r`nTextRange,4,IUIAutomationTextRange"
}

class UIA_TogglePattern extends UIA_Base
{
    ; https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationtogglepattern
    static __IID := "{94cf8058-9b8d-4ab9-8bfd-4cd0a33c8c70}"
	, __PatternID := 10015
	, __Properties := "CurrentToggleState,4,int`r`nCachedToggleState,5,int"
    Toggle() {
      return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
    }
}

class UIA_TransformPattern extends UIA_Base {
	static	__IID := "{A9B55844-A55D-4EF0-926D-569C16FF89BB}"
		,	__PatternID := 10016
		,	__Properties := "CurrentCanMove,6,bool`r`nCurrentCanResize,7,bool`r`nCurrentCanRotate,8,bool`r`nCachedCanMove,9,bool`r`nCachedCanResize,10,bool`r`nCachedCanRotate,11,bool"
	Move(x, y) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "double",x, "double",y))
	}
	Resize(w, h) {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "double",w, "double",h))
	}
	Rotate(degrees) {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "double",degrees))
	}
}

;~ class UIA_TransformPattern2 extends UIA_Base {10028
class UIA_TransformPattern2 extends UIA_TransformPattern { ; UNTESTED. It seems TransformPattern2 implementation doesn't extend TransformPattern methods/properties, yet TextPattern2 extends TextPattern... 
	static	__IID := "{6D74D017-6ECB-4381-B38B-3C17A48FF1C2}"
		,	__PatternID := 10028
		,	__Properties := UIA_TransformPattern.__Properties . "`r`nCurrentCanZoom,14,bool`r`nCachedCanZoom,15,bool`r`nCurrentZoomLevel,16,double`r`nCachedZoomLevel,17,double`r`nCurrentZoomMinimum,18,double`r`nCachedZoomMinimum,19,double`r`nCurrentZoomMaximum,20,double`r`nCachedZoomMaximum,21,double"
	Zoom(zoomValue) {
		return UIA_Hr(DllCall(this.__Vt(12), "ptr",this.__Value, "double",zoomValue))
	}
	ZoomByUnit(ZoomUnit) {
		return UIA_Hr(DllCall(this.__Vt(13), "ptr",this.__Value, "uint",ZoomUnit))
	}
}

class UIA_ValuePattern extends UIA_Base {
	static	__IID := "{A94CD8B1-0844-4CD6-9D2D-640537AB39E9}"
		,	__PatternID := 10002
		,	__Properties := "CurrentValue,4,BSTR`r`nCurrentIsReadOnly,5,BOOL`r`nCachedValue,6,double`r`nCachedIsReadOnly,7,BOOL"

	SetValue(val) {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr",&val))
	}
}

class UIA_VirtualizedItemPattern extends UIA_Base {
	static	__IID := "{6BA3D7A6-04CF-4F11-8793-A8D1CDE9969F}"
		,	__PatternID := 10020
		,	__Properties := ""

	Realize() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
	}
}

class UIA_WindowPattern extends UIA_Base {
	static __IID := "{0FAEF453-9208-43EF-BBB2-3B485177864F}"
		, __PatternID := 10009
		,  __Properties := "CurrentCanMaximize,6,int`r`nCurrentCanMinimize,7,int`r`nCurrentIsModal,8,int`r`nCurrentIsTopmost,9,int`r`nCurrentWindowVisualState,10,int`r`nCurrentWindowInteractionState,11,int`r`nCachedCanMaximize,12,int`r`nCachedCanMinimize,13,int`r`nCachedIsModal,14,int`r`nCachedIsTopmost,15,int`r`nCachedWindowVisualState,16,int`r`nCachedWindowInteractionState,17,int"
	Close() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value))
	}
	WaitForInputIdle(milliseconds) {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "int",milliseconds, "Int*", out))?out:
	}
	SetWindowVisualState(state) {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value, "int",state))
	}
}

class UIA_AnnotationPattern extends UIA_Base {	
	static __IID := "{9A175B21-339E-41B1-8E8B-623F6B681098}"
		, __PatternID := 10023
		,  __Properties := "CurrentAnnotationTypeId,3,int`r`nCurrentAnnotationTypeName,4,BSTR`r`nCurrentAuthor,5,BSTR`r`nCurrentDateTime,6,BSTR`r`nCurrentTarget,7,IUIAutomationElement`r`nCachedAnnotationTypeId,8,int`r`nCachedAnnotationTypeName,9,BSTR`r`nCachedAuthor,10,BSTR`r`nCachedDateTime,11,BSTR`r`nCachedTarget,12,IUIAutomationElement"
}

class UIA_DragPattern extends UIA_Base { ; UNTESTED, couldn't find a window that supported this
	static __IID := "{1DC7B570-1F54-4BAD-BCDA-D36A722FB7BD}"
		, __PatternID := 10030
		,  __Properties := "CurrentIsGrabbed,3,int`r`nCachedIsGrabbed,4,int`r`nCurrentDropEffect,5,BSTR`r`nCachedDropEffect,6,BSTR`r`nCurrentDropEffects,7,VARIANT`r`nCachedDropEffects,8,VARIANT`r`n"
	GetCurrentGrabbedItems() {
		return UIA_Hr(DllCall(this.__Vt(9), "ptr",this.__Value, "ptr*", out))?UIA_ElementArray(out):
	}
	GetCachedGrabbedItems() {
		return UIA_Hr(DllCall(this.__Vt(10), "ptr",this.__Value, "ptr*", out))?UIA_ElementArray(out):
	}
}

class UIA_DropTargetPattern extends UIA_Base { ; UNTESTED
	static __IID := "{69A095F7-EEE4-430E-A46B-FB73B1AE39A5}"
		, __PatternID := 10031
		,  __Properties := "CurrentDropTargetEffect,3,BSTR`r`nCachedDropTargetEffect,4,BSTR`r`nCurrentDropTargetEffects,5,VARIANT`r`nCachedDropTargetEffects,6,VARIANT`r`n"
}

class UIA_ObjectModelPattern extends UIA_Base {			; Windows 8 [desktop apps only]
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/hh437262(v=vs.85).aspx
	static	__IID := "{71c284b3-c14d-4d14-981e-19751b0d756d}"
		,	__PatternID := 10022
	
	GetUnderlyingObjectModel() { ; UNTESTED. Returns IUnknown interface used to access the underlying object model of the provider.
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr*", out))?out:
	}
}

;~ class UIA_PatternHandler extends UIA_Base {
;~ class UIA_PatternInstance extends UIA_Base {

class UIA_TextRange extends UIA_Base {
	static __IID := "{A543CC6A-F4AE-494B-8239-C814481187A8}"
		,  __Properties := ""
	
	Clone() {
		return UIA_Hr(DllCall(this.__Vt(3), "ptr",this.__Value, "ptr*",out))?UIA_TextRange(out):
	}
	Compare(comparisonTextRange) {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value,"ptr",comparisonTextRange.__Value, "ptr*",out))?out:
	}
	CompareEndPoints(srcEndPoint, comparisonTextRange, targetEndPoint) {
		return UIA_Hr(DllCall(this.__Vt(5), "ptr",this.__Value,"int", srcEndPoint,"ptr",comparisonTextRange.__Value, "int", targetEndPoint,"ptr*",out))?out:
	}
	ExpandToEnclosingUnit(unit=6) {
		return UIA_Hr(DllCall(this.__Vt(6), "ptr",this.__Value,"int",unit))
	}
	FindAttribute(attr, val, backward=False) {
		return UIA_Hr((A_PtrSize == 4) ? DllCall(this.__Vt(7), "ptr",this.__Value,"int",attr,"int64",NumGet(val, 0, "int64"),"int64",NumGet(val, 8, "int64"), "ptr*",out) : DllCall(this.__Vt(7), "ptr",this.__Value,"int",attr,"ptr",&val, "ptr*",out))?UIA_TextRange(out):
	}
	FindText(text, backward=False, ignoreCase=False) {
		return UIA_Hr(DllCall(this.__Vt(8), "ptr",this.__Value,"ptr", &text,"int",backward, "int", ignoreCase,"ptr*",out))?UIA_TextRange(out):
	}					
	GetAttributeValue(attr) {
		return UIA_Hr(DllCall(this.__Vt(9), "ptr",this.__Value,"int", attr,"ptr",UIA_Variant(out)))?UIA_VariantData(out):
	}
	GetBoundingRectangles() { ; Returns an array of rect objects {l:left,t:top,r:right,b:bottom}
		static b:={__Class:"object",__Type:"RECT",Struct:Func("UIA_RectStructure")}
		if UIA_Hr(DllCall(this.__Vt(10), "ptr",this.__Value,"ptr*",out)) {
			DllCall("oleaut32\SafeArrayGetVartype", "ptr", out, "ushort*", baseType)
			sa := UIA_GetSafeArrayValue(out, baseType), retArr := []
			Loop, % sa.MaxIndex() / 4
				retArr.Push({l:Floor(sa[4*(A_Index-1)+1]),t:Floor(sa[4*(A_Index-1)+2]),r:Floor(sa[4*(A_Index-1)+3]),b:Floor(sa[4*(A_Index-1)+4]),base:b})
			return retArr
		}
	}
	GetEnclosingElement() {
		return UIA_Hr(DllCall(this.__Vt(11), "ptr",this.__Value,"ptr*",out))?UIA_Element(out):
	}
	GetText(maxLength=-1) {
		return UIA_Hr(DllCall(this.__Vt(12), "ptr",this.__Value,"int", maxLength,"ptr*",out))?StrGet(out):
	}
	Move(unit, count) {
		return UIA_Hr(DllCall(this.__Vt(13), "ptr",this.__Value,"int", unit,"int",count, "ptr*",out))?out:
	}
	MoveEndpointByUnit(endpoint, unit, count) {
		return UIA_Hr(DllCall(this.__Vt(14), "ptr",this.__Value,"int", endpoint,"int", unit, "int", count, "ptr*",out))?out:
	}
	MoveEndpointByRange(srcEndPoint, range, targetEndPoint) {
		return UIA_Hr(DllCall(this.__Vt(15), "ptr",this.__Value,"int", srcEndPoint,"ptr",range.__Value, "int", targetEndPoint,"ptr*",out))
	}
	Select() {
		return UIA_Hr(DllCall(this.__Vt(16), "ptr",this.__Value))
	}
	AddToSelection() {
		return UIA_Hr(DllCall(this.__Vt(17), "ptr",this.__Value))
	}
	RemoveFromSelection() {
		return UIA_Hr(DllCall(this.__Vt(18), "ptr",this.__Value))
	}
	ScrollIntoView(alignToTop) {
		return UIA_Hr(DllCall(this.__Vt(19), "ptr",this.__Value,"int", alignToTop))
	}
	GetChildren() {
		return UIA_Hr(DllCall(this.__Vt(20), "ptr",this.__Value,"ptr*",out))?UIA_ElementArray(out):
	}
}

class UIA_TextRange2 extends UIA_TextRange {
	static __IID := "{BB9B40E0-5E04-46BD-9BE0-4B601B9AFAD4}"
		,  __Properties := ""
	ShowContextMenu() {
		return UIA_Hr(DllCall(this.__Vt(21), "ptr",this.__Value,"ptr*",out))
	}
}
class UIA_TextRange3 extends UIA_TextRange2 { ; UNTESTED
	static __IID := "{6A315D69-5512-4C2E-85F0-53FCE6DD4BC2}"
		,  __Properties := ""
	GetEnclosingElementBuildCache(cacheRequest) {
		return UIA_Hr(DllCall(this.__Vt(22), "Ptr", this.__Value, "Ptr", cacheRequest.__Value, "Ptr*", out))?UIA_Element(out):
	}
	GetChildrenBuildCache(cacheRequest) {
		return UIA_Hr(DllCall(this.__Vt(23), "Ptr", this.__Value, "Ptr", cacheRequest.__Value, "Ptr*", out))&&out?UIA_ElementArray(out):
	}
	GetAttributeValues(attributeIds, attributeIdCount) { ; currently returns a AHK array
		if ComObjValue(attributeIds)&0x2000
			SafeArray:=attributeIds
		else {
			SafeArray:=ComObj(0x2003,DllCall("oleaut32\SafeArrayCreateVector", "uint",13, "uint",0, "uint",attributeIds.MaxIndex()),1)
			for i,c in attributeIds
				SafeArray[A_Index-1]:=c.__Value, ObjAddRef(c.__Value) ; AddRef - SafeArrayDestroy will release UIA_Conditions - they also release themselves
		}
		return UIA_Hr(DllCall(this.__Vt(24), "ptr",this.__Value, "ptr", ComObjValue(SafeArray), "int", attributeIdCount, "ptr*",out))? UIA_SafeArrayToAHKArray(ComObj(0x2003,out,1)):
	}
}

class UIA_TextRangeArray extends UIA_Base {
	static __IID := "{CE4AE76A-E717-4C98-81EA-47371D028EB6}"
		,  __Properties := "Length,3,int"
	
	GetElement(i) {
		return UIA_Hr(DllCall(this.__Vt(4), "ptr",this.__Value, "int",i, "ptr*",out))?UIA_TextRange(out):
	}
}

{  ;~ UIA Functions
	UIA_Interface(maxVersion="") {
		static uia
		if IsObject(uia)
			return uia
		max := (maxVersion?maxVersion:UIA_Enum.UIA_MaxVersion_Interface)+1
		while (--max) {
			 
			if (!IsObject(UIA_Interface%max%) || (max == 1))
				continue

			try {
				if uia:=ComObjCreate("{e22ad333-b25f-460c-83d0-0581107395c9}",UIA_Interface%max%.__IID) {
					uia:=new UIA_Interface%max%(uia, 1, max), uiaBase := uia.base
					Loop, %max%
						uiaBase := uiaBase.base
					uiaBase.__UIA:=uia, uiaBase.TrueCondition:=uia.CreateTrueCondition(), uiaBase.TreeWalkerTrue := uia.CreateTreeWalker(uiaBase.TrueCondition)
					return uia
				}
			}
		}
		; If all else fails, try the first IUIAutomation version
		try {
			if uia:=ComObjCreate("{ff48dba4-60ef-4201-aa87-54103eef594e}","{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}")
				return uia:=new UIA_Interface(uia, 1, 1), uia.base.base.__UIA:=uia, uia.base.base.CurrentVersion:=1, uia.base.base.TrueCondition:=uia.CreateTrueCondition(), uia.base.base.TreeWalkerTrue := uia.CreateTreeWalker(uia.base.base.TrueCondition)
			throw "UIAutomation Interface failed to initialize."
		} catch e
			MsgBox, 262160, UIA Startup Error, % IsObject(e)?"IUIAutomation Interface is not registered.":e.Message
		return
	}
	UIA_Hr(hr) {
		;~ http://blogs.msdn.com/b/eldar/archive/2007/04/03/a-lot-of-hresult-codes.aspx
		static err:={0x8000FFFF:"Catastrophic failure.",0x80004001:"Not implemented.",0x8007000E:"Out of memory.",0x80070057:"One or more arguments are not valid.",0x80004002:"Interface not supported.",0x80004003:"Pointer not valid.",0x80070006:"Handle not valid.",0x80004004:"Operation aborted.",0x80004005:"Unspecified error.",0x80070005:"General access denied.",0x800401E5:"The object identified by this moniker could not be found.",0x80040201:"UIA_E_ELEMENTNOTAVAILABLE",0x80040200:"UIA_E_ELEMENTNOTENABLED",0x80131509:"UIA_E_INVALIDOPERATION",0x80040202:"UIA_E_NOCLICKABLEPOINT",0x80040204:"UIA_E_NOTSUPPORTED",0x80040203:"UIA_E_PROXYASSEMBLYNOTLOADED"} ; //not completed
		if hr&&(hr&=0xFFFFFFFF) {
			RegExMatch(Exception("",-2).what,"(\w+).(\w+)",i)
			throw Exception(UIA_Hex(hr) " - " err[hr], -2, i2 "  (" i1 ")")
		}
		return !hr
	}
	UIA_NotImplemented() {
		RegExMatch(Exception("",-2).What,"(\D+)\.(\D+)",m)
		MsgBox, 262192, UIA Message, Class:`t%m1%`nMember:`t%m2%`n`nMethod has not been implemented yet.
	}
	UIA_Element(e) {
		static v, previousVersion
		if (previousVersion != UIA_Enum.UIA_CurrentVersion_Element) ; Check if the user wants an element with a different version
			v := ""
		else if v
			return (v==1)?new UIA_Element(e,,1):new UIA_Element%v%(e,,v)
		max := UIA_Enum.UIA_CurrentVersion_Element+1
		While (--max) {
			if UIA_GUID(riid, UIA_Element%max%.__IID)
				return new UIA_Element%max%(e,,v:=max)
		}
		return new UIA_Element(e,,v:=1)
	}
	UIA_TextRange(e) {
		static v, previousVersion
		if (previousVersion != UIA_Enum.UIA_CurrentVersion_TextRange) ; Check if the user wants an element with a different version
			v := ""
		else if v
			return (v==1)?new UIA_TextRange(e,,1):new UIA_TextRange%v%(e,,v)
		max := UIA_Enum.UIA_MaxVersion_TextRange+1
		While (--max) {
			if UIA_GUID(riid, UIA_TextRange%max%.__IID)
				return new UIA_TextRange%max%(e,,v:=max)
		}
		return new UIA_TextRange(e,,v:=1)
	}
	UIA_Pattern(p) { ; Create a new pattern with the highest available version. This doesn't always result in good outcomes though: TransformPattern2 doesn't extend TransformPattern for some reason, yet TextPattern2 does extend TextPattern. Thus I left the current functionality of getting the highest available version, but when trying to call TransformPattern, specify UIA_Pattern("TransformPattern")
		static validated
		if !IsObject(validated)
			validated := {}
		patternName := InStr(p, "Pattern") ? p : p "Pattern", i:=1
		if (v:=validated[patternName])
			return (v==1)?patternName:(patternName v)
		if !(IsObject(UIA_%patternName%)&&UIA_%patternName%.__iid&&UIA_%patternName%.__PatternID) {
			return
		}
		Loop {
			i++
			if !(IsObject(UIA_%patternName%%i%)&&UIA_%patternName%%i%.__iid&&UIA_%patternName%%i%.__PatternID)
				break
		}
		if (--i > 1)
			return patternName i, validated[patternName] := i
		return patternName, validated[patternName] := 1
	}
	UIA_Enum(e) {
		if ObjHasKey(UIA_Enum, e)
			return UIA_Enum[e]
		else if ObjHasKey(UIA_Enum, "UIA_" e)
			return UIA_Enum["UIA_" e]
	}
	UIA_ElementArray(p, uia="") { ; should AHK Object be 0 or 1 based? /// answer: 1 based ///
		a:=new UIA_ElementArray(p),out:=[]
		Loop % a.Length
			out[A_Index]:=a.GetElement(A_Index-1)
		return out, out.base:={UIA_ElementArray:a}
	}
	UIA_TextRangeArray(p, uia="") { ; should AHK Object be 0 or 1 based? /// answer: 1 based ///
		a:=new UIA_TextRangeArray(p),out:=[]
		Loop % a.Length
			out[A_Index]:=a.GetElement(A_Index-1)
		return out, out.base:={UIA_TextRangeArray:a}
	}
	UIA_RectToObject(ByRef r) { ; rect.__Value work with DllCalls?
		static b:={__Class:"object",__Type:"RECT",Struct:Func("UIA_RectStructure")}
		return {l:NumGet(r,0,"Int"),t:NumGet(r,4,"Int"),r:NumGet(r,8,"Int"),b:NumGet(r,12,"Int"),base:b}
	}
	UIA_RectStructure(this, ByRef r) {
		static sides:="ltrb"
		VarSetCapacity(r,16)
		Loop Parse, sides
			NumPut(this[A_LoopField],r,(A_Index-1)*4,"Int")
	}
	UIA_SafeArrayToAHKArray(safearray) {
		out := []
		for k in safearray
			out.Push(k)
		return out
	}
	UIA_SafeArraysToObject(keys,values) {
	;~	1 dim safearrays w/ same # of elements
		out:={}
		for key in keys
			out[key]:=values[A_Index-1]
		return out
	}
	UIA_Hex(p) {
		setting:=A_FormatInteger
		SetFormat,IntegerFast,H
		out:=p+0 ""
		SetFormat,IntegerFast,%setting%
		return out
	}
	UIA_GUID(ByRef GUID, sGUID) { ;~ Converts a string to a binary GUID and returns its address.
		if !sGUID
			return
		VarSetCapacity(GUID,16,0)
		return DllCall("ole32\CLSIDFromString", "wstr",sGUID, "ptr",&GUID)>=0?&GUID:""
	}
	UIA_Variant(ByRef var,type=0,val=0) {
		; Does a variant need to be cleared? If it uses SysAllocString? 
		return (VarSetCapacity(var,8+2*A_PtrSize)+NumPut(type,var,0,"short")+NumPut(type=8? DllCall("oleaut32\SysAllocString", "ptr",&val):val,var,8,"ptr"))*0+&var
	}
	UIA_IsVariant(ByRef vt, ByRef type="", offset=0) {
		size:=VarSetCapacity(vt),type:=NumGet(vt,offset,"UShort")
		return size>=16&&size<=24&&type>=0&&(type<=23||type|0x2000)
	}
	UIA_VariantType(type){
		static _:={2:[2,"short"]
		,3:[4,"int"]
		,4:[4,"float"]
		,5:[8,"double"]
		,0xA:[4,"uint"]
		,0xB:[2,"short"]
		,0x10:[1,"char"]
		,0x11:[1,"uchar"]
		,0x12:[2,"ushort"]
		,0x13:[4,"uint"]
		,0x14:[8,"int64"]
		,0x15:[8,"uint64"]}
		return _.haskey(type)?_[type]:[A_PtrSize,"ptr"]
	}
	UIA_VariantData(ByRef p, flag=1, offset=0) {
		; based on Sean's COM_Enumerate function
		; need to clear varaint? what if you still need it (flag param)?
		return !UIA_IsVariant(p,vt, offset)?"Invalid Variant"
				:vt=0?"" ; VT_EMPTY
				:vt=3?NumGet(p,offset+8,"int")
				:vt=8?StrGet(NumGet(p,offset+8))
				:vt=11?-NumGet(p,offset+8,"short")
				:vt=9||vt=13||vt&0x2000?ComObj(vt,NumGet(p,offset+8),flag)
				:vt<0x1000&&UIA_VariantChangeType(&p,&p)=0?StrGet(NumGet(p,offset+8)) UIA_VariantClear(&p)
				:NumGet(p,offset+8)
	/*
		VT_EMPTY     =      0  		; No value
		VT_NULL      =      1 		; SQL-style Null
		VT_I2        =      2 		; 16-bit signed int
		VT_I4        =      3 		; 32-bit signed int
		VT_R4        =      4 		; 32-bit floating-point number
		VT_R8        =      5 		; 64-bit floating-point number
		VT_CY        =      6 		; Currency
		VT_DATE      =      7  		; Date
		VT_BSTR      =      8 		; COM string (Unicode string with length prefix)
		VT_DISPATCH  =      9 		; COM object 
		VT_ERROR     =    0xA  10	; Error code (32-bit integer)
		VT_BOOL      =    0xB  11	; Boolean True (-1) or False (0)
		VT_VARIANT   =    0xC  12	; VARIANT (must be combined with VT_ARRAY or VT_BYREF)
		VT_UNKNOWN   =    0xD  13	; IUnknown interface pointer
		VT_DECIMAL   =    0xE  14	; (not supported)
		VT_I1        =   0x10  16	; 8-bit signed int
		VT_UI1       =   0x11  17	; 8-bit unsigned int
		VT_UI2       =   0x12  18	; 16-bit unsigned int
		VT_UI4       =   0x13  19	; 32-bit unsigned int
		VT_I8        =   0x14  20	; 64-bit signed int
		VT_UI8       =   0x15  21	; 64-bit unsigned int
		VT_INT       =   0x16  22	; Signed machine int
		VT_UINT      =   0x17  23	; Unsigned machine int
		VT_RECORD    =   0x24  36	; User-defined type
		VT_ARRAY     = 0x2000  		; SAFEARRAY
		VT_BYREF     = 0x4000  		; Pointer to another type of value
					 = 0x1000  4096

		COM_VariantChangeType(pvarDst, pvarSrc, vt=8) {
			return DllCall("oleaut32\VariantChangeTypeEx", "ptr",pvarDst, "ptr",pvarSrc, "Uint",1024, "Ushort",0, "Ushort",vt)
		}
		COM_VariantClear(pvar) {
			DllCall("oleaut32\VariantClear", "ptr",pvar)
		}
		COM_SysAllocString(str) {
			Return	DllCall("oleaut32\SysAllocString", "Uint", &str)
		}
		COM_SysFreeString(pstr) {
				DllCall("oleaut32\SysFreeString", "Uint", pstr)
		}
		COM_SysString(ByRef wString, sString) {
			VarSetCapacity(wString,4+nLen:=2*StrLen(sString))
			Return	DllCall("kernel32\lstrcpyW","Uint",NumPut(nLen,wString),"Uint",&sString)
		}
		DllCall("oleaut32\SafeArrayGetVartype", "ptr*",ComObjValue(SafeArray), "uint*",pvt)
		HRESULT SafeArrayGetVartype(
		  _In_   SAFEARRAY *psa,
		  _Out_  VARTYPE *pvt
		);

		DllCall("oleaut32\SafeArrayDestroy", "ptr",ComObjValue(SafeArray))
		HRESULT SafeArrayDestroy(
		  _In_  SAFEARRAY *psa
		);
	*/
	}
	UIA_VariantChangeType(pvarDst, pvarSrc, vt=8) { ; written by Sean
		return DllCall("oleaut32\VariantChangeTypeEx", "ptr",pvarDst, "ptr",pvarSrc, "Uint",1024, "Ushort",0, "Ushort",vt)
	}
	UIA_VariantClear(pvar) { ; Written by Sean
		DllCall("oleaut32\VariantClear", "ptr",pvar)
	}
	UIA_GetSafeArrayValue(p,type){ ; Credit: https://github.com/neptercn/UIAutomation/blob/master/UIA.ahk
		t:=UIA_VariantType(type),item:={},pv:=NumGet(p+8+A_PtrSize,"ptr")
		loop % NumGet(p+8+2*A_PtrSize,"uint") {
			item.Insert((type=8)?StrGet(NumGet(pv+(A_Index-1)*t.1,t.2),"utf-16"):NumGet(pv+(A_Index-1)*t.1,t.2))
		}
		return item
	}
	UIA_CreateEventHandler(funcName, handlerType="") { ; Possible handlerType values: empty, FocusChanged, StructureChanged, TextEditTextChanged, Changes, Notification
		if !IsFunc(funcName){
			msgbox %funcName% is not a function.
			return
		}
		static ptr
		VarSetCapacity(ptr,A_PtrSize*5)
		handler := new _UIA_%handlerType%EventHandler(&ptr,1,funcName), ObjAddRef(ptr) ; deref will be done on destruction of EventHandler. Function name piggybacks on the __Version property
		,NumPut(&ptr+A_PtrSize,ptr)
		,NumPut(RegisterCallback("_UIA_QueryInterface","F"),ptr,A_PtrSize*1)
		,NumPut(RegisterCallback("_UIA_AddRef","F"),ptr,A_PtrSize*2)
		,NumPut(RegisterCallback("_UIA_Release","F"),ptr,A_PtrSize*3)
		,NumPut(RegisterCallback("_UIA_" handlerType "EventHandler.Handle" (handlerType == "" ? "Automation" : handlerType) "Event","F",,&handler),ptr,A_PtrSize*4)
		return handler	
	}
	_UIA_QueryInterface(pSelf, pRIID, pObj){ ; Credit: https://github.com/neptercn/UIAutomation/blob/master/UIA.ahk
		DllCall("ole32\StringFromIID","ptr",pRIID,"ptr*",sz),str:=StrGet(sz)
		return (str="{00000000-0000-0000-C000-000000000046}")||(str="{146c3c17-f12e-4e22-8c27-f894b9b79c69}")||(str="{40cd37d4-c756-4b0c-8c6f-bddfeeb13b50}")||(str="{e81d1b4e-11c5-42f8-9754-e7036c79f054}")||(str="{c270f6b5-5c69-4290-9745-7a7f97169468}")?NumPut(pSelf,pObj+0)*0:0x80004002 ; E_NOINTERFACE
	}
	_UIA_AddRef(pSelf){
	}
	_UIA_Release(pSelf){
	}
}










/*
	UIA_Enum contains UIA constants and enumerations. 
	
	There are multiple ways to access the constants:
	1) Use the returned object of UIA_Interface:
		After UIA := UIA_Interface()
		UIA.UIA_InvokePatternId would return the corresponding ID of 10000. Similarly UIA.TreeScope_Descendants would return 0x4
		If a property starts with "UIA_" then that part may be omitted: UIA.ButtonControlTypeId would return 50000.
		
		Calling UIA_Enum methods is also supported:
		UIA.UIA_EventId(20000) would return "ToolTipOpened"
	2) Use UIA_Enum prototype. This requires using exact property and method names:
		UIA_Enum.UIA_InvokePatternId
		UIA_Enum.UIA_PatternId(10000)
	
	3) Use a UIA_Enum object:
		After UIAc := new UIA_Enum
		UIAc.InvokePatternId returns 10000
		UIAc.PatternId(10000) returns "InvokePattern"
	
	4) Use the UIA_Enum function:
		UIA_Enum("ButtonControlTypeId") would return 50000
*/

class UIA_Enum { ; main source: https://github.com/Ixiko/AHK-libs-and-classes-collection/blob/master/libs/o-z/UIAutomationClient_1_0_64bit.ahk

	__Get(member) {
		if member not in base 
		{
			if ((SubStr(member, 1, 4) != "UIA_") && ObjHasKey(UIA_Enum, "UIA_" member)) {
				return this["UIA_" member]
			}
		}
		return UIA_Enum.member
	}
	__Call(member, params*) {
		if member not in base 
		{
			if ((SubStr(member, 1, 4) != "UIA_") && IsFunc("UIA_Enum.UIA_" member))
				return this["UIA_" member](params*)
		}
	}

	; UIA_Interface specific enums, which define maximum available version numbers for interfaces (eg currently the highest version for UIA_Interface is 7). CurrentVersion specifies the version used by any new objects created, this can be changed dynamically.
	static UIA_MaxVersion_Interface := 7
	static UIA_MaxVersion_Element := 7
	static UIA_CurrentVersion_Element := 7
	static UIA_MaxVersion_TextRange := 3
	static UIA_CurrentVersion_TextRange := 3
	
	; The following are not strictly enumerations but constants, it makes more sense to include them here
	; module UIA_PatternIds
	static UIA_InvokePatternId := 10000
	static UIA_SelectionPatternId := 10001
	static UIA_ValuePatternId := 10002
	static UIA_RangeValuePatternId := 10003
	static UIA_ScrollPatternId := 10004
	static UIA_ExpandCollapsePatternId := 10005
	static UIA_GridPatternId := 10006
	static UIA_GridItemPatternId := 10007
	static UIA_MultipleViewPatternId := 10008
	static UIA_WindowPatternId := 10009
	static UIA_SelectionItemPatternId := 10010
	static UIA_DockPatternId := 10011
	static UIA_TablePatternId := 10012
	static UIA_TableItemPatternId := 10013
	static UIA_TextPatternId := 10014
	static UIA_TogglePatternId := 10015
	static UIA_TransformPatternId := 10016
	static UIA_ScrollItemPatternId := 10017
	static UIA_LegacyIAccessiblePatternId := 10018
	static UIA_ItemContainerPatternId := 10019
	static UIA_VirtualizedItemPatternId := 10020
	static UIA_SynchronizedInputPatternId := 10021
	static UIA_ObjectModelPatternId := 10022
	static UIA_AnnotationPatternId := 10023
	static UIA_TextPattern2Id := 10024
	static UIA_StylesPatternId := 10025
	static UIA_SpreadsheetPatternId := 10026
	static UIA_SpreadsheetItemPatternId := 10027
	static UIA_TransformPattern2Id := 10028
	static UIA_TextChildPatternId := 10029
	static UIA_DragPatternId := 10030
	static UIA_DropTargetPatternId := 10031
	static UIA_TextEditPatternId := 10032
	static UIA_CustomNavigationPatternId := 10033
	static UIA_SelectionPattern2Id := 10034 

	UIA_PatternId(n="") {
		static name:={10000:"InvokePattern",10001:"SelectionPattern",10002:"ValuePattern",10003:"RangeValuePattern",10004:"ScrollPattern",10005:"ExpandCollapsePattern",10006:"GridPattern",10007:"GridItemPattern",10008:"MultipleViewPattern",10009:"WindowPattern",10010:"SelectionItemPattern",10011:"DockPattern",10012:"TablePattern",10013:"TableItemPattern",10014:"TextPattern",10015:"TogglePattern",10016:"TransformPattern",10017:"ScrollItemPattern",10018:"LegacyIAccessiblePattern",10019:"ItemContainerPattern",10020:"VirtualizedItemPattern",10021:"SynchronizedInputPattern",10022:"ObjectModelPattern",10023:"AnnotationPattern",10024:"TextPattern2",10025:"StylesPattern",10026:"SpreadsheetPattern",10027:"SpreadsheetItemPattern",10028:"TransformPattern2",10029:"TextChildPattern",10030:"DragPattern",10031:"DropTargetPattern",10032:"TextEditPattern",10033:"CustomNavigationPattern",10034:"SelectionPattern2"}, id:={InvokePattern:10000,SelectionPattern:10001,ValuePattern:10002,RangeValuePattern:10003,ScrollPattern:10004,ExpandCollapsePattern:10005,GridPattern:10006,GridItemPattern:10007,MultipleViewPattern:10008,WindowPattern:10009,SelectionItemPattern:10010,DockPattern:10011,TablePattern:10012,TableItemPattern:10013,TextPattern:10014,TogglePattern:10015,TransformPattern:10016,ScrollItemPattern:10017,LegacyIAccessiblePattern:10018,ItemContainerPattern:10019,VirtualizedItemPattern:10020,SynchronizedInputPattern:10021,ObjectModelPattern:10022,AnnotationPattern:10023,TextPattern2:10024,StylesPattern:10025,SpreadsheetPattern:10026,SpreadsheetItemPattern:10027,TransformPattern2:10028,TextChildPattern:10029,DragPattern:10030,DropTargetPattern:10031,TextEditPattern:10032,CustomNavigationPattern:10033,SelectionPattern2:10034}
		if !n
			return id
		if n is integer
			return name[n]
		if ObjHasKey(id, n "Pattern")
			return id[n "Pattern"]
		else if ObjHasKey(id, n)
			return id[n]
		return id[RegexReplace(n, "(?:UIA_)?(.+?)(?:Id)?$", "$1")]
	}

	; module UIA_EventIds
	static UIA_ToolTipOpenedEventId := 20000
	static UIA_ToolTipClosedEventId := 20001
	static UIA_StructureChangedEventId := 20002
	static UIA_MenuOpenedEventId := 20003
	static UIA_AutomationPropertyChangedEventId := 20004
	static UIA_AutomationFocusChangedEventId := 20005
	static UIA_AsyncContentLoadedEventId := 20006
	static UIA_MenuClosedEventId := 20007
	static UIA_LayoutInvalidatedEventId := 20008
	static UIA_Invoke_InvokedEventId := 20009
	static UIA_SelectionItem_ElementAddedToSelectionEventId := 20010
	static UIA_SelectionItem_ElementRemovedFromSelectionEventId := 20011
	static UIA_SelectionItem_ElementSelectedEventId := 20012
	static UIA_Selection_InvalidatedEventId := 20013
	static UIA_Text_TextSelectionChangedEventId := 20014
	static UIA_Text_TextChangedEventId := 20015
	static UIA_Window_WindowOpenedEventId := 20016
	static UIA_Window_WindowClosedEventId := 20017
	static UIA_MenuModeStartEventId := 20018
	static UIA_MenuModeEndEventId := 20019
	static UIA_InputReachedTargetEventId := 20020
	static UIA_InputReachedOtherElementEventId := 20021
	static UIA_InputDiscardedEventId := 20022
	static UIA_SystemAlertEventId := 20023
	static UIA_LiveRegionChangedEventId := 20024
	static UIA_HostedFragmentRootsInvalidatedEventId := 20025
	static UIA_Drag_DragStartEventId := 20026
	static UIA_Drag_DragCancelEventId := 20027
	static UIA_Drag_DragCompleteEventId := 20028
	static UIA_DropTarget_DragEnterEventId := 20029
	static UIA_DropTarget_DragLeaveEventId := 20030
	static UIA_DropTarget_DroppedEventId := 20031
	static UIA_TextEdit_TextChangedEventId := 20032
	static UIA_TextEdit_ConversionTargetChangedEventId := 20033
	static UIA_ChangesEventId := 20034
	static UIA_NotificationEventId := 20035
	static UIA_ActiveTextPositionChangedEventId := 20036

	UIA_EventId(n="") {
		static id:={ToolTipOpened:20000,ToolTipClosed:20001,StructureChanged:20002,MenuOpened:20003,AutomationPropertyChanged:20004,AutomationFocusChanged:20005,AsyncContentLoaded:20006,MenuClosed:20007,LayoutInvalidated:20008,Invoke_Invoked:20009,SelectionItem_ElementAddedToSelection:20010,SelectionItem_ElementRemovedFromSelection:20011,SelectionItem_ElementSelected:20012,Selection_Invalidated:20013,Text_TextSelectionChanged:20014,Text_TextChanged:20015,Window_WindowOpened:20016,Window_WindowClosed:20017,MenuModeStart:20018,MenuModeEnd:20019,InputReachedTarget:20020,InputReachedOtherElement:20021,InputDiscarded:20022,SystemAlert:20023,LiveRegionChanged:20024,HostedFragmentRootsInvalidated:20025,Drag_DragStart:20026,Drag_DragCancel:20027,Drag_DragComplete:20028,DropTarget_DragEnter:20029,DropTarget_DragLeave:20030,DropTarget_Dropped:20031,TextEdit_TextChanged:20032,TextEdit_ConversionTargetChanged:20033,Changes:20034,Notification:20035,ActiveTextPositionChanged:20036}, name:={20000:"ToolTipOpened",20001:"ToolTipClosed",20002:"StructureChanged",20003:"MenuOpened",20004:"AutomationPropertyChanged",20005:"AutomationFocusChanged",20006:"AsyncContentLoaded",20007:"MenuClosed",20008:"LayoutInvalidated",20009:"Invoke_Invoked",20010:"SelectionItem_ElementAddedToSelection",20011:"SelectionItem_ElementRemovedFromSelection",20012:"SelectionItem_ElementSelected",20013:"Selection_Invalidated",20014:"Text_TextSelectionChanged",20015:"Text_TextChanged",20016:"Window_WindowOpened",20017:"Window_WindowClosed",20018:"MenuModeStart",20019:"MenuModeEnd",20020:"InputReachedTarget",20021:"InputReachedOtherElement",20022:"InputDiscarded",20023:"SystemAlert",20024:"LiveRegionChanged",20025:"HostedFragmentRootsInvalidated",20026:"Drag_DragStart",20027:"Drag_DragCancel",20028:"Drag_DragComplete",20029:"DropTarget_DragEnter",20030:"DropTarget_DragLeave",20031:"DropTarget_Dropped",20032:"TextEdit_TextChanged",20033:"TextEdit_ConversionTargetChanged",20034:"Changes",20035:"Notification",20036:"ActiveTextPositionChanged"}
		if !n
			return id
		if n is integer
			return name[n]
		if ObjHasKey(id, n)
			return id[n]
		return id[StrReplace(StrReplace(n, "EventId"), "UIA_")]
	}

	; module UIA_PropertyIds
	static UIA_RuntimeIdPropertyId := 30000
	static UIA_BoundingRectanglePropertyId := 30001
	static UIA_ProcessIdPropertyId := 30002
	static UIA_ControlTypePropertyId := 30003
	static UIA_LocalizedControlTypePropertyId := 30004
	static UIA_NamePropertyId := 30005
	static UIA_AcceleratorKeyPropertyId := 30006
	static UIA_AccessKeyPropertyId := 30007
	static UIA_HasKeyboardFocusPropertyId := 30008
	static UIA_IsKeyboardFocusablePropertyId := 30009
	static UIA_IsEnabledPropertyId := 30010
	static UIA_AutomationIdPropertyId := 30011
	static UIA_ClassNamePropertyId := 30012
	static UIA_HelpTextPropertyId := 30013
	static UIA_ClickablePointPropertyId := 30014
	static UIA_CulturePropertyId := 30015
	static UIA_IsControlElementPropertyId := 30016
	static UIA_IsContentElementPropertyId := 30017
	static UIA_LabeledByPropertyId := 30018
	static UIA_IsPasswordPropertyId := 30019
	static UIA_NativeWindowHandlePropertyId := 30020
	static UIA_ItemTypePropertyId := 30021
	static UIA_IsOffscreenPropertyId := 30022
	static UIA_OrientationPropertyId := 30023
	static UIA_FrameworkIdPropertyId := 30024
	static UIA_IsRequiredForFormPropertyId := 30025
	static UIA_ItemStatusPropertyId := 30026
	static UIA_IsDockPatternAvailablePropertyId := 30027
	static UIA_IsExpandCollapsePatternAvailablePropertyId := 30028
	static UIA_IsGridItemPatternAvailablePropertyId := 30029
	static UIA_IsGridPatternAvailablePropertyId := 30030
	static UIA_IsInvokePatternAvailablePropertyId := 30031
	static UIA_IsMultipleViewPatternAvailablePropertyId := 30032
	static UIA_IsRangeValuePatternAvailablePropertyId := 30033
	static UIA_IsScrollPatternAvailablePropertyId := 30034
	static UIA_IsScrollItemPatternAvailablePropertyId := 30035
	static UIA_IsSelectionItemPatternAvailablePropertyId := 30036
	static UIA_IsSelectionPatternAvailablePropertyId := 30037
	static UIA_IsTablePatternAvailablePropertyId := 30038
	static UIA_IsTableItemPatternAvailablePropertyId := 30039
	static UIA_IsTextPatternAvailablePropertyId := 30040
	static UIA_IsTogglePatternAvailablePropertyId := 30041
	static UIA_IsTransformPatternAvailablePropertyId := 30042
	static UIA_IsValuePatternAvailablePropertyId := 30043
	static UIA_IsWindowPatternAvailablePropertyId := 30044
	static UIA_ValueValuePropertyId := 30045
	static UIA_ValueIsReadOnlyPropertyId := 30046
	static UIA_RangeValueValuePropertyId := 30047
	static UIA_RangeValueIsReadOnlyPropertyId := 30048
	static UIA_RangeValueMinimumPropertyId := 30049
	static UIA_RangeValueMaximumPropertyId := 30050
	static UIA_RangeValueLargeChangePropertyId := 30051
	static UIA_RangeValueSmallChangePropertyId := 30052
	static UIA_ScrollHorizontalScrollPercentPropertyId := 30053
	static UIA_ScrollHorizontalViewSizePropertyId := 30054
	static UIA_ScrollVerticalScrollPercentPropertyId := 30055
	static UIA_ScrollVerticalViewSizePropertyId := 30056
	static UIA_ScrollHorizontallyScrollablePropertyId := 30057
	static UIA_ScrollVerticallyScrollablePropertyId := 30058
	static UIA_SelectionSelectionPropertyId := 30059
	static UIA_SelectionCanSelectMultiplePropertyId := 30060
	static UIA_SelectionIsSelectionRequiredPropertyId := 30061
	static UIA_GridRowCountPropertyId := 30062
	static UIA_GridColumnCountPropertyId := 30063
	static UIA_GridItemRowPropertyId := 30064
	static UIA_GridItemColumnPropertyId := 30065
	static UIA_GridItemRowSpanPropertyId := 30066
	static UIA_GridItemColumnSpanPropertyId := 30067
	static UIA_GridItemContainingGridPropertyId := 30068
	static UIA_DockDockPositionPropertyId := 30069
	static UIA_ExpandCollapseExpandCollapseStatePropertyId := 30070
	static UIA_MultipleViewCurrentViewPropertyId := 30071
	static UIA_MultipleViewSupportedViewsPropertyId := 30072
	static UIA_WindowCanMaximizePropertyId := 30073
	static UIA_WindowCanMinimizePropertyId := 30074
	static UIA_WindowWindowVisualStatePropertyId := 30075
	static UIA_WindowWindowInteractionStatePropertyId := 30076
	static UIA_WindowIsModalPropertyId := 30077
	static UIA_WindowIsTopmostPropertyId := 30078
	static UIA_SelectionItemIsSelectedPropertyId := 30079
	static UIA_SelectionItemSelectionContainerPropertyId := 30080
	static UIA_TableRowHeadersPropertyId := 30081
	static UIA_TableColumnHeadersPropertyId := 30082
	static UIA_TableRowOrColumnMajorPropertyId := 30083
	static UIA_TableItemRowHeaderItemsPropertyId := 30084
	static UIA_TableItemColumnHeaderItemsPropertyId := 30085
	static UIA_ToggleToggleStatePropertyId := 30086
	static UIA_TransformCanMovePropertyId := 30087
	static UIA_TransformCanResizePropertyId := 30088
	static UIA_TransformCanRotatePropertyId := 30089
	static UIA_IsLegacyIAccessiblePatternAvailablePropertyId := 30090
	static UIA_LegacyIAccessibleChildIdPropertyId := 30091
	static UIA_LegacyIAccessibleNamePropertyId := 30092
	static UIA_LegacyIAccessibleValuePropertyId := 30093
	static UIA_LegacyIAccessibleDescriptionPropertyId := 30094
	static UIA_LegacyIAccessibleRolePropertyId := 30095
	static UIA_LegacyIAccessibleStatePropertyId := 30096
	static UIA_LegacyIAccessibleHelpPropertyId := 30097
	static UIA_LegacyIAccessibleKeyboardShortcutPropertyId := 30098
	static UIA_LegacyIAccessibleSelectionPropertyId := 30099
	static UIA_LegacyIAccessibleDefaultActionPropertyId := 30100
	static UIA_AriaRolePropertyId := 30101
	static UIA_AriaPropertiesPropertyId := 30102
	static UIA_IsDataValidForFormPropertyId := 30103
	static UIA_ControllerForPropertyId := 30104
	static UIA_DescribedByPropertyId := 30105
	static UIA_FlowsToPropertyId := 30106
	static UIA_ProviderDescriptionPropertyId := 30107
	static UIA_IsItemContainerPatternAvailablePropertyId := 30108
	static UIA_IsVirtualizedItemPatternAvailablePropertyId := 30109
	static UIA_IsSynchronizedInputPatternAvailablePropertyId := 30110
	static UIA_OptimizeForVisualContentPropertyId := 30111
	static UIA_IsObjectModelPatternAvailablePropertyId := 30112
	static UIA_AnnotationAnnotationTypeIdPropertyId := 30113
	static UIA_AnnotationAnnotationTypeNamePropertyId := 30114
	static UIA_AnnotationAuthorPropertyId := 30115
	static UIA_AnnotationDateTimePropertyId := 30116
	static UIA_AnnotationTargetPropertyId := 30117
	static UIA_IsAnnotationPatternAvailablePropertyId := 30118
	static UIA_IsTextPattern2AvailablePropertyId := 30119
	static UIA_StylesStyleIdPropertyId := 30120
	static UIA_StylesStyleNamePropertyId := 30121
	static UIA_StylesFillColorPropertyId := 30122
	static UIA_StylesFillPatternStylePropertyId := 30123
	static UIA_StylesShapePropertyId := 30124
	static UIA_StylesFillPatternColorPropertyId := 30125
	static UIA_StylesExtendedPropertiesPropertyId := 30126
	static UIA_IsStylesPatternAvailablePropertyId := 30127
	static UIA_IsSpreadsheetPatternAvailablePropertyId := 30128
	static UIA_SpreadsheetItemFormulaPropertyId := 30129
	static UIA_SpreadsheetItemAnnotationObjectsPropertyId := 30130
	static UIA_SpreadsheetItemAnnotationTypesPropertyId := 30131
	static UIA_IsSpreadsheetItemPatternAvailablePropertyId := 30132
	static UIA_Transform2CanZoomPropertyId := 30133
	static UIA_IsTransformPattern2AvailablePropertyId := 30134
	static UIA_LiveSettingPropertyId := 30135
	static UIA_IsTextChildPatternAvailablePropertyId := 30136
	static UIA_IsDragPatternAvailablePropertyId := 30137
	static UIA_DragIsGrabbedPropertyId := 30138
	static UIA_DragDropEffectPropertyId := 30139
	static UIA_DragDropEffectsPropertyId := 30140
	static UIA_IsDropTargetPatternAvailablePropertyId := 30141
	static UIA_DropTargetDropTargetEffectPropertyId := 30142
	static UIA_DropTargetDropTargetEffectsPropertyId := 30143
	static UIA_DragGrabbedItemsPropertyId := 30144
	static UIA_Transform2ZoomLevelPropertyId := 30145
	static UIA_Transform2ZoomMinimumPropertyId := 30146
	static UIA_Transform2ZoomMaximumPropertyId := 30147
	static UIA_FlowsFromPropertyId := 30148
	static UIA_IsTextEditPatternAvailablePropertyId := 30149
	static UIA_IsPeripheralPropertyId := 30150
	static UIA_IsCustomNavigationPatternAvailablePropertyId := 30151
	static UIA_PositionInSetPropertyId := 30152
	static UIA_SizeOfSetPropertyId := 30153
	static UIA_LevelPropertyId := 30154
	static UIA_AnnotationTypesPropertyId := 30155
	static UIA_AnnotationObjectsPropertyId := 30156
	static UIA_LandmarkTypePropertyId := 30157
	static UIA_LocalizedLandmarkTypePropertyId := 30158
	static UIA_FullDescriptionPropertyId := 30159
	static UIA_FillColorPropertyId := 30160
	static UIA_OutlineColorPropertyId := 30161
	static UIA_FillTypePropertyId := 30162
	static UIA_VisualEffectsPropertyId := 30163
	static UIA_OutlineThicknessPropertyId := 30164
	static UIA_CenterPointPropertyId := 30165
	static UIA_RotationPropertyId := 30166
	static UIA_SizePropertyId := 30167
	static UIA_IsSelectionPattern2AvailablePropertyId := 30168
	static UIA_Selection2FirstSelectedItemPropertyId := 30169
	static UIA_Selection2LastSelectedItemPropertyId := 30170
	static UIA_Selection2CurrentSelectedItemPropertyId := 30171
	static UIA_Selection2ItemCountPropertyId := 30172
	static UIA_HeadingLevelPropertyId := 30173
	static UIA_IsDialogPropertyId := 30174

	UIA_PropertyId(n="") {
		static ids:="RuntimeId:30000,BoundingRectangle:30001,ProcessId:30002,ControlType:30003,LocalizedControlType:30004,Name:30005,AcceleratorKey:30006,AccessKey:30007,HasKeyboardFocus:30008,IsKeyboardFocusable:30009,IsEnabled:30010,AutomationId:30011,ClassName:30012,HelpText:30013,ClickablePoint:30014,Culture:30015,IsControlElement:30016,IsContentElement:30017,LabeledBy:30018,IsPassword:30019,NativeWindowHandle:30020,ItemType:30021,IsOffscreen:30022,Orientation:30023,FrameworkId:30024,IsRequiredForForm:30025,ItemStatus:30026,IsDockPatternAvailable:30027,IsExpandCollapsePatternAvailable:30028,IsGridItemPatternAvailable:30029,IsGridPatternAvailable:30030,IsInvokePatternAvailable:30031,IsMultipleViewPatternAvailable:30032,IsRangeValuePatternAvailable:30033,IsScrollPatternAvailable:30034,IsScrollItemPatternAvailable:30035,IsSelectionItemPatternAvailable:30036,IsSelectionPatternAvailable:30037,IsTablePatternAvailable:30038,IsTableItemPatternAvailable:30039,IsTextPatternAvailable:30040,IsTogglePatternAvailable:30041,IsTransformPatternAvailable:30042,IsValuePatternAvailable:30043,IsWindowPatternAvailable:30044,ValueValue:30045,ValueIsReadOnly:30046,RangeValueValue:30047,RangeValueIsReadOnly:30048,RangeValueMinimum:30049,RangeValueMaximum:30050,RangeValueLargeChange:30051,RangeValueSmallChange:30052,ScrollHorizontalScrollPercent:30053,ScrollHorizontalViewSize:30054,ScrollVerticalScrollPercent:30055,ScrollVerticalViewSize:30056,ScrollHorizontallyScrollable:30057,ScrollVerticallyScrollable:30058,SelectionSelection:30059,SelectionCanSelectMultiple:30060,SelectionIsSelectionRequired:30061,GridRowCount:30062,GridColumnCount:30063,GridItemRow:30064,GridItemColumn:30065,GridItemRowSpan:30066,GridItemColumnSpan:30067,GridItemContainingGrid:30068,DockDockPosition:30069,ExpandCollapseExpandCollapseState:30070,MultipleViewCurrentView:30071,MultipleViewSupportedViews:30072,WindowCanMaximize:30073,WindowCanMinimize:30074,WindowWindowVisualState:30075,WindowWindowInteractionState:30076,WindowIsModal:30077,WindowIsTopmost:30078,SelectionItemIsSelected:30079,SelectionItemSelectionContainer:30080,TableRowHeaders:30081,TableColumnHeaders:30082,TableRowOrColumnMajor:30083,TableItemRowHeaderItems:30084,TableItemColumnHeaderItems:30085,ToggleToggleState:30086,TransformCanMove:30087,TransformCanResize:30088,TransformCanRotate:30089,IsLegacyIAccessiblePatternAvailable:30090,LegacyIAccessibleChildId:30091,LegacyIAccessibleName:30092,LegacyIAccessibleValue:30093,LegacyIAccessibleDescription:30094,LegacyIAccessibleRole:30095,LegacyIAccessibleState:30096,LegacyIAccessibleHelp:30097,LegacyIAccessibleKeyboardShortcut:30098,LegacyIAccessibleSelection:30099,LegacyIAccessibleDefaultAction:30100,AriaRole:30101,AriaProperties:30102,IsDataValidForForm:30103,ControllerFor:30104,DescribedBy:30105,FlowsTo:30106,ProviderDescription:30107,IsItemContainerPatternAvailable:30108,IsVirtualizedItemPatternAvailable:30109,IsSynchronizedInputPatternAvailable:30110,OptimizeForVisualContent:30111,IsObjectModelPatternAvailable:30112,AnnotationAnnotationTypeId:30113,AnnotationAnnotationTypeName:30114,AnnotationAuthor:30115,AnnotationDateTime:30116,AnnotationTarget:30117,IsAnnotationPatternAvailable:30118,IsTextPattern2Available:30119,StylesStyleId:30120,StylesStyleName:30121,StylesFillColor:30122,StylesFillPatternStyle:30123,StylesShape:30124,StylesFillPatternColor:30125,StylesExtendedProperties:30126,IsStylesPatternAvailable:30127,IsSpreadsheetPatternAvailable:30128,SpreadsheetItemFormula:30129,SpreadsheetItemAnnotationObjects:30130,SpreadsheetItemAnnotationTypes:30131,IsSpreadsheetItemPatternAvailable:30132,Transform2CanZoom:30133,IsTransformPattern2Available:30134,LiveSetting:30135,IsTextChildPatternAvailable:30136,IsDragPatternAvailable:30137,DragIsGrabbed:30138,DragDropEffect:30139,DragDropEffects:30140,IsDropTargetPatternAvailable:30141,DropTargetDropTargetEffect:30142,DropTargetDropTargetEffects:30143,DragGrabbedItems:30144,Transform2ZoomLevel:30145,Transform2ZoomMinimum:30146,Transform2ZoomMaximum:30147,FlowsFrom:30148,IsTextEditPatternAvailable:30149,IsPeripheral:30150,IsCustomNavigationPatternAvailable:30151,PositionInSet:30152,SizeOfSet:30153,Level:30154,AnnotationTypes:30155,AnnotationObjects:30156,LandmarkType:30157,LocalizedLandmarkType:30158,FullDescription:30159,FillColor:30160,OutlineColor:30161,FillType:30162,VisualEffects:30163,OutlineThickness:30164,CenterPoint:30165,Rotation:30166,Size:30167,IsSelectionPattern2Available:30168,Selection2FirstSelectedItem:30169,Selection2LastSelectedItem:30170,Selection2CurrentSelectedItem:30171,Selection2ItemCount:30173,IsDialog:30174"
		if !n
			return ids		
		if n is integer 
		{
			RegexMatch(ids, "([^,]+):" n, m)
			return m1
		}
		
		n := StrReplace(StrReplace(n, "UIA_"), "PropertyId")
		RegexMatch(ids, "(?:^|,)" n "(?:" n ")?(?:Id)?:(\d+)", m)
		return m1
	}

	UIA_PropertyVariantType(id){
		static type:={30000:0x2003,30001:0x2005,30002:3,30003:3,30004:8,30005:8,30006:8,30007:8,30008:0xB,30009:0xB,30010:0xB,30011:8,30012:8,30013:8,30014:0x2005,30015:3,30016:0xB,30017:0xB,30018:0xD,30019:0xB,30020:3,30021:8,30022:0xB,30023:3,30024:8,30025:0xB,30026:8,30027:0xB,30028:0xB,30029:0xB,30030:0xB,30031:0xB,30032:0xB,30033:0xB,30034:0xB,30035:0xB,30036:0xB,30037:0xB,30038:0xB,30039:0xB,30040:0xB,30041:0xB,30042:0xB,30043:0xB,30044:0xB,30045:8,30046:0xB,30047:5,30048:0xB,30049:5,30050:5,30051:5,30052:5,30053:5,30054:5,30055:5,30056:5,30057:0xB,30058:0xB,30059:0x200D,30060:0xB,30061:0xB,30062:3,30063:3,30064:3,30065:3,30066:3,30067:3,30068:0xD,30069:3,30070:3,30071:3,30072:0x2003,30073:0xB,30074:0xB,30075:3,30076:3,30077:0xB,30078:0xB,30079:0xB,30080:0xD,30081:0x200D,30082:0x200D,30083:0x2003,30084:0x200D,30085:0x200D,30086:3,30087:0xB,30088:0xB,30089:0xB,30090:0xB,30091:3,30092:8,30093:8,30094:8,30095:3,30096:3,30097:8,30098:8,30099:0x200D,30100:8}, type2:={30101:8,30102:8,30103:0xB,30104:0xD,30105:0xD,30106:0xD,30107:8,30108:0xB,30109:0xB,30110:0xB,30111:0xB,30112:0xB,30113:3,30114:8,30115:8,30116:8,30117:0xD,30118:0xB,30119:0xB,30120:3,30121:8,30122:3,30123:8,30124:8,30125:3,30126:8,30127:0xB,30128:0xB,30129:8,30130:0x200D,30131:0x2003,30132:0xB,30133:0xB,30134:0xB,30135:3,30136:0xB,30137:0xB,30138:0xB,30139:8,30140:0x2008,30141:0xB,30142:8,30143:0x2008,30144:0x200D,30145:5,30146:5,30147:5,30148:0x200D,30149:0xB,30150:0xB,30151:0xB,30152:3,30153:3,30154:3,30155:0x2003,30156:0x2003,30157:3,30158:8,30159:8,30160:3,30161:0x2003,30162:3,30163:3,30164:0x2005,30165:0x2005,30166:5,30167:0x2005,30168:0xB} ; missing VTs from 30169 and on
		return ObjHasKey(type, id) ? type[id] : type2[id]
	}

	; module UIA_TextAttributeIds
	static UIA_AnimationStyleAttributeId := 40000
	static UIA_BackgroundColorAttributeId := 40001
	static UIA_BulletStyleAttributeId := 40002
	static UIA_CapStyleAttributeId := 40003
	static UIA_CultureAttributeId := 40004
	static UIA_FontNameAttributeId := 40005
	static UIA_FontSizeAttributeId := 40006
	static UIA_FontWeightAttributeId := 40007
	static UIA_ForegroundColorAttributeId := 40008
	static UIA_HorizontalTextAlignmentAttributeId := 40009
	static UIA_IndentationFirstLineAttributeId := 40010
	static UIA_IndentationLeadingAttributeId := 40011
	static UIA_IndentationTrailingAttributeId := 40012
	static UIA_IsHiddenAttributeId := 40013
	static UIA_IsItalicAttributeId := 40014
	static UIA_IsReadOnlyAttributeId := 40015
	static UIA_IsSubscriptAttributeId := 40016
	static UIA_IsSuperscriptAttributeId := 40017
	static UIA_MarginBottomAttributeId := 40018
	static UIA_MarginLeadingAttributeId := 40019
	static UIA_MarginTopAttributeId := 40020
	static UIA_MarginTrailingAttributeId := 40021
	static UIA_OutlineStylesAttributeId := 40022
	static UIA_OverlineColorAttributeId := 40023
	static UIA_OverlineStyleAttributeId := 40024
	static UIA_StrikethroughColorAttributeId := 40025
	static UIA_StrikethroughStyleAttributeId := 40026
	static UIA_TabsAttributeId := 40027
	static UIA_TextFlowDirectionsAttributeId := 40028
	static UIA_UnderlineColorAttributeId := 40029
	static UIA_UnderlineStyleAttributeId := 40030
	static UIA_AnnotationTypesAttributeId := 40031
	static UIA_AnnotationObjectsAttributeId := 40032
	static UIA_StyleNameAttributeId := 40033
	static UIA_StyleIdAttributeId := 40034
	static UIA_LinkAttributeId := 40035
	static UIA_IsActiveAttributeId := 40036
	static UIA_SelectionActiveEndAttributeId := 40037
	static UIA_CaretPositionAttributeId := 40038
	static UIA_CaretBidiModeAttributeId := 40039
	static UIA_LineSpacingAttributeId := 40040
	static UIA_BeforeParagraphSpacingAttributeId := 40041
	static UIA_AfterParagraphSpacingAttributeId := 40042
	static UIA_SayAsInterpretAsAttributeId := 40043

	UIA_AttributeId(n="") {
		static id:={AnimationStyle:40000,BackgroundColor:40001,BulletStyle:40002,CapStyle:40003,Culture:40004,FontName:40005,FontSize:40006,FontWeight:40007,ForegroundColor:40008,HorizontalTextAlignment:40009,IndentationFirstLine:40010,IndentationLeading:40011,IndentationTrailing:40012,IsHidden:40013,IsItalic:40014,IsReadOnly:40015,IsSubscript:40016,IsSuperscript:40017,MarginBottom:40018,MarginLeading:40019,MarginTop:40020,MarginTrailing:40021,OutlineStyles:40022,OverlineColor:40023,OverlineStyle:40024,StrikethroughColor:40025,StrikethroughStyle:40026,Tabs:40027,TextFlowDirections:40028,UnderlineColor:40029,UnderlineStyle:40030,AnnotationTypes:40031,AnnotationObjects:40032,StyleName:40033,StyleId:40034,Link:40035,IsActive:40036,SelectionActiveEnd:40037,CaretPosition:40038,CaretBidiMode:40039,LineSpacing:40040,BeforeParagraphSpacing:40041,AfterParagraphSpacing:40042,SayAsInterpretAs:40043}, name:={40000:"AnimationStyle",40001:"BackgroundColor",40002:"BulletStyle",40003:"CapStyle",40004:"Culture",40005:"FontName",40006:"FontSize",40007:"FontWeight",40008:"ForegroundColor",40009:"HorizontalTextAlignment",40010:"IndentationFirstLine",40011:"IndentationLeading",40012:"IndentationTrailing",40013:"IsHidden",40014:"IsItalic",40015:"IsReadOnly",40016:"IsSubscript",40017:"IsSuperscript",40018:"MarginBottom",40019:"MarginLeading",40020:"MarginTop",40021:"MarginTrailing",40022:"OutlineStyles",40023:"OverlineColor",40024:"OverlineStyle",40025:"StrikethroughColor",40026:"StrikethroughStyle",40027:"Tabs",40028:"TextFlowDirections",40029:"UnderlineColor",40030:"UnderlineStyle",40031:"AnnotationTypes",40032:"AnnotationObjects",40033:"StyleName",40034:"StyleId",40035:"Link",40036:"IsActive",40037:"SelectionActiveEnd",40038:"CaretPosition",40039:"CaretBidiMode",40040:"LineSpacing",40041:"BeforeParagraphSpacing",40042:"AfterParagraphSpacing",40043:"SayAsInterpretAs"}
		if !n
			return id
		if n is integer
			return name[n]
		if ObjHasKey(id, n)
			return id[n]
		return id[StrReplace(StrReplace(n, "AttributeId"), "UIA_")]
	}

	UIA_AttributeVariantType(id){
		Static type:={40000:3,40001:3,40002:3,40003:3,40004:3,40005:8,40006:5,40007:3,40008:3,40009:3,40010:5,40011:5,40012:5,40013:0xB,40014:0xB,40015:0xB,40016:0xB,40017:0xB,40018:5,40019:5,40020:5,40021:5,40022:3,40023:3,40024:3,40025:3,40026:3,40027:0x2005,40028:3,40029:3,40030:3,40031:0x2003,40032:0x200D,40033:8,40034:3,40035:0xD,40036:0xB,40037:3,40038:3,40039:3,40040:8,40041:5,40042:5,40043:8} ; 40032 and 40043 might be incorrect
		return type[id]
	}

	; module UIA_ControlTypeIds
	static UIA_ButtonControlTypeId := 50000
	static UIA_CalendarControlTypeId := 50001
	static UIA_CheckBoxControlTypeId := 50002
	static UIA_ComboBoxControlTypeId := 50003
	static UIA_EditControlTypeId := 50004
	static UIA_HyperlinkControlTypeId := 50005
	static UIA_ImageControlTypeId := 50006
	static UIA_ListItemControlTypeId := 50007
	static UIA_ListControlTypeId := 50008
	static UIA_MenuControlTypeId := 50009
	static UIA_MenuBarControlTypeId := 50010
	static UIA_MenuItemControlTypeId := 50011
	static UIA_ProgressBarControlTypeId := 50012
	static UIA_RadioButtonControlTypeId := 50013
	static UIA_ScrollBarControlTypeId := 50014
	static UIA_SliderControlTypeId := 50015
	static UIA_SpinnerControlTypeId := 50016
	static UIA_StatusBarControlTypeId := 50017
	static UIA_TabControlTypeId := 50018
	static UIA_TabItemControlTypeId := 50019
	static UIA_TextControlTypeId := 50020
	static UIA_ToolBarControlTypeId := 50021
	static UIA_ToolTipControlTypeId := 50022
	static UIA_TreeControlTypeId := 50023
	static UIA_TreeItemControlTypeId := 50024
	static UIA_CustomControlTypeId := 50025
	static UIA_GroupControlTypeId := 50026
	static UIA_ThumbControlTypeId := 50027
	static UIA_DataGridControlTypeId := 50028
	static UIA_DataItemControlTypeId := 50029
	static UIA_DocumentControlTypeId := 50030
	static UIA_SplitButtonControlTypeId := 50031
	static UIA_WindowControlTypeId := 50032
	static UIA_PaneControlTypeId := 50033
	static UIA_HeaderControlTypeId := 50034
	static UIA_HeaderItemControlTypeId := 50035
	static UIA_TableControlTypeId := 50036
	static UIA_TitleBarControlTypeId := 50037
	static UIA_SeparatorControlTypeId := 50038
	static UIA_SemanticZoomControlTypeId := 50039
	static UIA_AppBarControlTypeId := 50040

	UIA_ControlTypeId(n="") {
		static id:={Button:50000,Calendar:50001,CheckBox:50002,ComboBox:50003,Edit:50004,Hyperlink:50005,Image:50006,ListItem:50007,List:50008,Menu:50009,MenuBar:50010,MenuItem:50011,ProgressBar:50012,RadioButton:50013,ScrollBar:50014,Slider:50015,Spinner:50016,StatusBar:50017,Tab:50018,TabItem:50019,Text:50020,ToolBar:50021,ToolTip:50022,Tree:50023,TreeItem:50024,Custom:50025,Group:50026,Thumb:50027,DataGrid:50028,DataItem:50029,Document:50030,SplitButton:50031,Window:50032,Pane:50033,Header:50034,HeaderItem:50035,Table:50036,TitleBar:50037,Separator:50038,SemanticZoom:50039,AppBar:50040}, name:={50000:"Button",50001:"Calendar",50002:"CheckBox",50003:"ComboBox",50004:"Edit",50005:"Hyperlink",50006:"Image",50007:"ListItem",50008:"List",50009:"Menu",50010:"MenuBar",50011:"MenuItem",50012:"ProgressBar",50013:"RadioButton",50014:"ScrollBar",50015:"Slider",50016:"Spinner",50017:"StatusBar",50018:"Tab",50019:"TabItem",50020:"Text",50021:"ToolBar",50022:"ToolTip",50023:"Tree",50024:"TreeItem",50025:"Custom",50026:"Group",50027:"Thumb",50028:"DataGrid",50029:"DataItem",50030:"Document",50031:"SplitButton",50032:"Window",50033:"Pane",50034:"Header",50035:"HeaderItem",50036:"Table",50037:"TitleBar",50038:"Separator",50039:"SemanticZoom",50040:"AppBar"}
		if !n
			return id
		if n is integer
			return name[n]
		if ObjHasKey(id, n)
			return id[n]
		return id[StrReplace(StrReplace(n, "ControlTypeId"), "UIA_")]
	}

; module AnnotationType
	static UIA_AnnotationType_Unknown := 60000
	static UIA_AnnotationType_SpellingError := 60001
	static UIA_AnnotationType_GrammarError := 60002
	static UIA_AnnotationType_Comment := 60003
	static UIA_AnnotationType_FormulaError := 60004
	static UIA_AnnotationType_TrackChanges := 60005
	static UIA_AnnotationType_Header := 60006
	static UIA_AnnotationType_Footer := 60007
	static UIA_AnnotationType_Highlighted := 60008
	static UIA_AnnotationType_Endnote := 60009
	static UIA_AnnotationType_Footnote := 60010
	static UIA_AnnotationType_InsertionChange := 60011
	static UIA_AnnotationType_DeletionChange := 60012
	static UIA_AnnotationType_MoveChange := 60013
	static UIA_AnnotationType_FormatChange := 60014
	static UIA_AnnotationType_UnsyncedChange := 60015
	static UIA_AnnotationType_EditingLockedChange := 60016
	static UIA_AnnotationType_ExternalChange := 60017
	static UIA_AnnotationType_ConflictingChange := 60018
	static UIA_AnnotationType_Author := 60019
	static UIA_AnnotationType_AdvancedProofingIssue := 60020
	static UIA_AnnotationType_DataValidationError := 60021
	static UIA_AnnotationType_CircularReferenceError := 60022
	static UIA_AnnotationType_Mathematics := 60023

	UIA_AnnotationType(n="") {
		static id:={Unknown:60000,SpellingError:60001,GrammarError:60002,Comment:60003,FormulaError:60004,TrackChanges:60005,Header:60006,Footer:60007,Highlighted:60008,Endnote:60009,Footnote:60010,InsertionChange:60011,DeletionChange:60012,MoveChange:60013,FormatChange:60014,UnsyncedChange:60015,EditingLockedChange:60016,ExternalChange:60017,ConflictingChange:60018,Author:60019,AdvancedProofingIssue:60020,DataValidationError:60021,CircularReferenceError:60022,Mathematics:60023}, name:={60000:"Unknown",60001:"SpellingError",60002:"GrammarError",60003:"Comment",60004:"FormulaError",60005:"TrackChanges",60006:"Header",60007:"Footer",60008:"Highlighted",60009:"Endnote",60010:"Footnote",60011:"InsertionChange",60012:"DeletionChange",60013:"MoveChange",60014:"FormatChange",60015:"UnsyncedChange",60016:"EditingLockedChange",60017:"ExternalChange",60018:"ConflictingChange",60019:"Author",60020:"AdvancedProofingIssue",60021:"DataValidationError",60022:"CircularReferenceError",60023:"Mathematics"}
		if !n
			return id
		if n is integer
			return name[n]
		if ObjHasKey(id, n)
			return id[n]
		return id[StrSplit(n, "_").Pop()]
	}

	; module StyleId
	static UIA_StyleId_Custom := 70000
	static UIA_StyleId_Heading1 := 70001
	static UIA_StyleId_Heading2 := 70002
	static UIA_StyleId_Heading3 := 70003
	static UIA_StyleId_Heading4 := 70004
	static UIA_StyleId_Heading5 := 70005
	static UIA_StyleId_Heading6 := 70006
	static UIA_StyleId_Heading7 := 70007
	static UIA_StyleId_Heading8 := 70008
	static UIA_StyleId_Heading9 := 70009
	static UIA_StyleId_Title := 70010
	static UIA_StyleId_Subtitle := 70011
	static UIA_StyleId_Normal := 70012
	static UIA_StyleId_Emphasis := 70013
	static UIA_StyleId_Quote := 70014
	static UIA_StyleId_BulletedList := 70015
	static UIA_StyleId_NumberedList := 70016

	UIA_StyleId(n="") {
		static id:={Custom:70000,Heading1:70001,Heading2:70002,Heading3:70003,Heading4:70004,Heading5:70005,Heading6:70006,Heading7:70007,Heading8:70008,Heading9:70009,Title:70010,Subtitle:70011,Normal:70012,Emphasis:70013,Quote:70014,BulletedList:70015,NumberedList:70016}, name:={70000:"Custom",70001:"Heading1",70002:"Heading2",70003:"Heading3",70004:"Heading4",70005:"Heading5",70006:"Heading6",70007:"Heading7",70008:"Heading8",70009:"Heading9",70010:"Title",70011:"Subtitle",70012:"Normal",70013:"Emphasis",70014:"Quote",70015:"BulletedList",70016:"NumberedList"}
		if !n
			return id
		if n is integer
			return name[n]
		if ObjHasKey(id, n)
			return id[n]
		return id[StrSplit(n, "_").Pop()]
	}

	; module LandmarkTypeIds
	static UIA_CustomLandmarkTypeId := 80000
	static UIA_FormLandmarkTypeId := 80001
	static UIA_MainLandmarkTypeId := 80002
	static UIA_NavigationLandmarkTypeId := 80003
	static UIA_SearchLandmarkTypeId := 80004
	
	UIA_LandmarkTypeId(n="") {
		static id:={Custom:80000,Form:80001,Main:80002,Navigation:80003,Search:80004}, name:={80000:"Custom",80001:"Form",80002:"Main",80003:"Navigation",80004:"Search"}
		if !n
			return id
		if n is integer
			return name[n]
		if ObjHasKey(id, n)
			return id[n]
		return id[StrReplace(StrReplace(n, "LandmarkTypeId"), "UIA_")]
	}
 
	; module HeadingLevelIds
	static UIA_HeadingLevel_None := 80050
	static UIA_HeadingLevel1 := 80051
	static UIA_HeadingLevel2 := 80052
	static UIA_HeadingLevel3 := 80053
	static UIA_HeadingLevel4 := 80054
	static UIA_HeadingLevel5 := 80055
	static UIA_HeadingLevel6 := 80056
	static UIA_HeadingLevel7 := 80057
	static UIA_HeadingLevel8 := 80058
	static UIA_HeadingLevel9 := 80059

	UIA_HeadingLevel(n="") {
		static id:={None:80050, 1:80051, 2:80052, 3:80053, 4:80054, 5:80055, 6:80056, 7:80057, 8:80058, 9:80059}, name:={80050:"None", 80051:"1", 80052:"2", 80053:"3", 80054:"4", 80055:"5", 80056:"6", 80057:"7", 80058:"8", 80059:"9"}
		if !n
			return id
		if n is integer
			return (n > 80000)?name[n]:id[n]
		if ObjHasKey(id, n)
			return id[n]
		return id[StrReplace(StrReplace(n, "HeadingLevel"), "UIA_")]
	}

	; module ChangeIds
	static UIA_SummaryChangeId := 90000
	
	UIA_ChangeId(n="") {
		static id:={Summary:90000}, name:={90000:"Summary"}
		if !n
			return id
		if n is integer
			return name[n]
		if ObjHasKey(id, n)
			return id[n]
		return id[StrReplace(StrReplace(n, "ChangeId"), "UIA_")]
	}

	; module MetadataIds
	static UIA_SayAsInterpretAsMetadataId := 100000
	
	UIA_MetadataId(n="") {
		static id:={SayAsInterpretAs:100000}, name:={100000:"SayAsInterpretAs"}
		if !n
			return id
		if n is integer
			return name[n]
		if ObjHasKey(id, n)
			return id[n]
		return id[StrReplace(StrReplace(n, "MetadataId"), "UIA_")]
	}


	; Uiautomationcoreapi.h enumerations
	; enum AsyncContentLoadedState Contains values that describe the progress of asynchronous loading of content.
	static AsyncContentLoadedState_Beginning := 0
	static AsyncContentLoadedState_Progress := 1
	static AsyncContentLoadedState_Completed := 2
	
	AsyncContentLoadedState(Value="") {
		static v1:={0:"Beginning", 1:"Progress", 2:"Completed"}
		if Value is not integer
			return this["AsyncContentLoadedState_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum AutomationIdentifierType Contains values used in the UiaLookupId function
	static AutomationIdentifierType_Property := 0
	static AutomationIdentifierType_Pattern := 1
	static AutomationIdentifierType_Event := 2
	static AutomationIdentifierType_ControlType := 3
	static AutomationIdentifierType_TextAttribute := 4
	static AutomationIdentifierType_LandmarkType := 5
	static AutomationIdentifierType_Annotation := 6
	static AutomationIdentifierType_Changes := 7
	static AutomationIdentifierType_Style := 8

	AutomationIdentifierType(Value="") {
		static v1:={0:"Property", 1:"Pattern", 2:"Event", 3:"ControlType", 4:"TextAttribute", 5:"LandmarkType", 6:"Annotation", 7:"Changes", 8:"Style"}
		if Value is not integer
			return this["AutomationIdentifierType_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum ConditionType Contains values that specify a type of UiaCondition.
	static ConditionType_True := 0
	static ConditionType_False := 1
	static ConditionType_Property := 2
	static ConditionType_And := 3
	static ConditionType_Or := 4
	static ConditionType_Not := 5

	ConditionType(Value="") {
		static v1:={0:"True", 1:"False", 2:"Property", 3:"And", 4:"Or", 5:"Not"}
		if Value is not integer
			return this["ConditionType_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum EventArgsType
	static EventArgsType_Simple := 0
	static EventArgsType_PropertyChanged := 1
	static EventArgsType_StructureChanged := 2
	static EventArgsType_AsyncContentLoaded := 3
	static EventArgsType_WindowClosed := 4
	static EventArgsType_TextEditTextChanged := 5
	static EventArgsType_Changes := 6
	static EventArgsType_Notification := 7
	static EventArgsType_ActiveTextPositionChanged := 8
	static EventArgsType_StructuredMarkup := 9

	EventArgsType(Value="") {
		static v1:={0:"Simple", 1:"PropertyChanged", 2:"StructureChanged", 3:"AsyncContentLoaded", 4:"WindowClosed", 5:"TextEditTextChanged", 6:"Changes", 7:"Notification", 8:"ActiveTextPositionChanged", 9:"StructuredMarkup"}
		if Value is not integer
			return this["EventArgsType_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; Uiautomationclient.h enumerations
	
	; enum AutomationElementMode Contains values that specify the type of reference to use when returning UI Automation elements.
	static AutomationElementMode_None := 0x0
	static AutomationElementMode_Full := 0x1

	AutomationElementMode(Value="") {
		static v1:={0x0:"None", 0x1:"Full"}
		if Value is not integer
			return this["AutomationElementMode_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum CoalesceEventsOptions Contains possible values for the CoalesceEvents property, which indicates whether an accessible technology client receives all events, or a subset where duplicate events are detected and filtered.
	static CoalesceEventsOptions_Disabled := 0x0
	static CoalesceEventsOptions_Enabled := 0x1

	CoalesceEventsOptions(Value="") {
		static v1:={0x0:"Disabled", 0x1:"Enabled"}
		if Value is not integer
			return this["CoalesceEventsOptions_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum ConnectionRecoveryBehaviorOptions Contains possible values for the ConnectionRecoveryBehavior property, which indicates whether an accessible technology client adjusts provider request timeouts when the provider is non-responsive.
	static ConnectionRecoveryBehaviorOptions_Disabled := 0
	static ConnectionRecoveryBehaviorOptions_Enabled := 0x1

	ConnectionRecoveryBehaviorOptions(Value="") {
		static v1:={0x0:"Disabled", 0x1:"Enabled"}
		if Value is not integer
			return this["ConnectionRecoveryBehaviorOptions_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum PropertyConditionFlags Contains values used in creating property conditions.
	static PropertyConditionFlags_None := 0x0
	static PropertyConditionFlags_IgnoreCase := 0x1

	PropertyConditionFlags(Value="") {
		static v1:={0x0:"None", 0x1:"IgnoreCase"}
		if Value is not integer
			return this["PropertyConditionFlags_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum TreeScope Contains values that specify the scope of various operations in the Microsoft UI Automation tree.
	static TreeScope_None := 0x0
	static TreeScope_Element := 0x1
	static TreeScope_Children := 0x2
	static TreeScope_Descendants := 0x4
	static TreeScope_Parent := 0x8
	static TreeScope_Ancestors := 0x10
	static TreeScope_Subtree := 0x7

	TreeScope(Value="") {
		static v1:={0x0:"None", 0x1:"Element", 0x2:"Children", 0x4:"Descendants", 0x8:"Parent", 0x10:"Ancestors", 0x7:"Subtree"}
		if Value is not integer
			return this["TreeScope_" Value]
		return (Value=="")?v1:v1[Value]
	}

	;enum TreeTraversalOptions Defines values that can be used to customize tree navigation order.
	static TreeTraversalOptions_Default := 0x0
	static TreeTraversalOptions_PostOrder := 0x1
	static TreeTraversalOptions_LastToFirstOrder := 0x2

	TreeTraversalOptions(Value="") {
		static v1:={0x0:"Default", 0x1:"PostOrder", 0x2:"LastToFirstOrder"}
		if Value is not integer
			return this["TreeTraversalOptions_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; Uiautomationcore.h enumerations

	; enum ActiveEnd Contains possible values for the SelectionActiveEnd text attribute, which indicates the location of the caret relative to a text range that represents the currently selected text.
	static ActiveEnd_None := 0
	static ActiveEnd_Start := 1
	static ActiveEnd_End := 2

	ActiveEnd(Value="") {
		static v1:={0x0:"None", 0x1:"Start", 0x2:"End"}
		if Value is not integer
			return this["ActiveEnd_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum AnimationStyle Contains values for the AnimationStyle text attribute.
	static AnimationStyle_None := 0
	static AnimationStyle_LasVegasLights := 1
	static AnimationStyle_BlinkingBackground := 2
	static AnimationStyle_SparkleText := 3
	static AnimationStyle_MarchingBlackAnts := 4
	static AnimationStyle_MarchingRedAnts := 5
	static AnimationStyle_Shimmer := 6
	static AnimationStyle_Other := -1

	AnimationStyle(Value="") {
		static v1:={0:"None", 1:"LasVegasLights",2:"BlinkingBackground", 3:"SparkleText",4:"MarchingBlackAnts", 5:"MarchingRedAnts", 6:"Shimmer", -1:"Other"}
		if Value is not integer
			return this["AnimationStyle_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum BulletStyle Contains values for the BulletStyle text attribute.
	static BulletStyle_None := 0
	static BulletStyle_HollowRoundBullet := 1
	static BulletStyle_FilledRoundBullet := 2
	static BulletStyle_HollowSquareBullet := 3
	static BulletStyle_FilledSquareBullet := 4
	static BulletStyle_DashBullet := 5
	static BulletStyle_Other := -1

	BulletStyle(Value="") {
		static v1:={0:"None", 1:"HollowRoundBullet",2:"FilledRoundBullet", 3:"HollowSquareBullet",4:"FilledSquareBullet", 5:"DashBullet", -1:"Other"}
		if Value is not integer
			return this["BulletStyle_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum CapStyle Contains values that specify the value of the CapStyle text attribute.
	static CapStyle_None := 0
	static CapStyle_SmallCap := 1
	static CapStyle_AllCap := 2
	static CapStyle_AllPetiteCaps := 3
	static CapStyle_PetiteCaps := 4
	static CapStyle_Unicase := 5
	static CapStyle_Titling := 6
	static CapStyle_Other := -1
	
	CapStyle(Value="") {
		static v1:={0:"None", 1:"SmallCap",2:"AllCap", 3:"AllPetiteCaps",4:"PetiteCaps", 5:"Unicase",6:"Titling", -1:"Other"}
		if Value is not integer
			return this["CapStyle_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum CaretBidiMode Contains possible values for the CaretBidiMode text attribute, which indicates whether the caret is in text that flows from left to right, or from right to left.
	static CaretBidiMode_LTR := 0
	static CaretBidiMode_RTL := 1

	CaretBidiMode(Value="") {
		static v1:={0:"LTR", 1:"RTL"}
		if Value is not integer
			return this["CaretBidiMode_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum CaretPosition Contains possible values for the CaretPosition text attribute, which indicates the location of the caret relative to a line of text in a text range.
	static CaretPosition_Unknown := 0
	static CaretPosition_EndOfLine := 1
	static CaretPosition_BeginningOfLine := 2
	
	CaretPosition(Value="") {
		static v1:={0:"Unknown", 1:"EndOfLine", 2:"BeginningOfLine"}
		if Value is not integer
			return this["CaretPosition_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum DockPosition Contains values that specify the location of a docking window represented by the Dock control pattern.
	static DockPosition_Top := 0x0
	static DockPosition_Left := 0x1
	static DockPosition_Bottom := 0x2
	static DockPosition_Right := 0x3
	static DockPosition_Fill := 0x4
	static DockPosition_None := 0x5

	DockPosition(Value="") {
		static v1:={0x0:"Top", 0x1:"Left", 0x2:"Bottom", 0x3:"Right", 0x4:"Fill", 0x5:"None"}
		if Value is not integer
			return this["DockPosition_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum ExpandCollapseState Contains values that specify the state of a UI element that can be expanded and collapsed.	
	static ExpandCollapseState_Collapsed := 0x0
	static ExpandCollapseState_Expanded := 0x1
	static ExpandCollapseState_PartiallyExpanded := 0x2
	static ExpandCollapseState_LeafNode := 0x3

	ExpandCollapseState(Value="") {
		static v1:={0x0:"Collapsed", 0x1:"Expanded", 0x2:"PartiallyExpanded", 0x3:"LeafNode"}
		if Value is not integer
			return this["ExpandCollapseState_" Value]
		return (Value=="")?v1:v1[Value]
	}	
	
	; enum FillType Contains values for the FillType attribute.
	static FillType_None := 0
	static FillType_Color := 1
	static FillType_Gradient := 2
	static FillType_Picture := 3
	static FillType_Pattern := 4

	FillType(Value="") {
		static v1:={0x0:"None", 0x1:"Color", 0x2:"Gradient", 0x3:"Picture", 0x4:"Pattern"}
		if Value is not integer
			return this["FillType_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum FlowDirection Contains values for the TextFlowDirections text attribute.
	static FlowDirections_Default := 0
	static FlowDirections_RightToLeft := 0x1
	static FlowDirections_BottomToTop := 0x2
	static FlowDirections_Vertical := 0x4

	FlowDirection(Value="") {
		static v1:={0x0:"Default", 0x1:"RightToLeft", 0x2:"BottomToTop", 0x4:"Vertical"}
		if Value is not integer
			return this["FlowDirection_" Value]
		return (Value=="")?v1:v1[Value]
	}

	;enum LiveSetting Contains possible values for the LiveSetting property. This property is implemented by provider elements that are part of a live region.
	static LiveSetting_Off := 0x0
	static LiveSetting_Polite := 0x1
	static LiveSetting_Assertive := 0x2

	LiveSetting(Value="") {
		static v1:={0x0:"Off", 0x1:"Polite", 0x2:"Assertive"}
		if Value is not integer
			return this["LiveSetting_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum NavigateDirection Contains values used to specify the direction of navigation within the Microsoft UI Automation tree.
	static NavigateDirection_Parent := 0x0
	static NavigateDirection_NextSibling := 0x1
	static NavigateDirection_PreviousSibling := 0x2
	static NavigateDirection_FirstChild := 0x3
	static NavigateDirection_LastChild := 0x4

	NavigateDirection(Value="") {
		static v1:={0x0:"Parent", 0x1:"NextSibling", 0x2:"PreviousSibling", 0x3:"FirstChild", 0x4:"LastChild"}
		if Value is not integer
			return this["NavigateDirection_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum NotificationKind Defines values that indicate the type of a notification event, and a hint to the listener about the processing of the event. 
	static NotificationKind_ItemAdded := 0
	static NotificationKind_ItemRemoved := 1
	static NotificationKind_ActionCompleted := 2
	static NotificationKind_ActionAborted := 3
	static NotificationKind_Other := 4
	
	NotificationKind(Value="") {
		static v1:={0x0:"ItemAdded", 0x1:"ItemRemoved", 0x2:"ActionCompleted", 0x3:"ActionAborted", 0x4:"Other"}
		if Value is not integer
			return this["NotificationKind_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum NotificationProcessing Defines values that indicate how a notification should be processed.
	static NotificationProcessing_ImportantAll := 0
	static NotificationProcessing_ImportantMostRecent := 1
	static NotificationProcessing_All := 2
	static NotificationProcessing_MostRecent := 3
	static NotificationProcessing_CurrentThenMostRecent := 4

	NotificationProcessing(Value="") {
		static v1:={0x0:"ImportantAll", 0x1:"ImportantMostRecent", 0x2:"All", 0x3:"MostRecent", 0x4:"CurrentThenMostRecent"}
		if Value is not integer
			return this["NotificationProcessing_" Value]
		return (Value=="")?v1:v1[Value]
	}	

	; enum OrientationType Contains values that specify the orientation of a control.
	static OrientationType_None := 0x0
	static OrientationType_Horizontal := 0x1
	static OrientationType_Vertical := 0x2

	OrientationType(Value="") {
		static v1:={0x0:"None", 0x1:"Horizontal", 0x2:"Vertical"}
		if Value is not integer
			return this["OrientationType_" Value]
		return (Value=="")?v1:v1[Value]
	}
	
	; enum OutlineStyles Contains values for the OutlineStyle text attribute.
	static OutlineStyles_None := 0
	static OutlineStyles_Outline := 1
	static OutlineStyles_Shadow := 2
	static OutlineStyles_Engraved := 4
	static OutlineStyles_Embossed := 8
	
	OutlineStyles(Value="") {
		static v1:={0:"None", 1:"Outline", 2:"Shadow", 4:"Engraved", 8:"Embossed"}
		if Value is not integer
			return this["OutlineStyles_" Value]
		return (Value=="")?v1:v1[Value]
	}

	;enum ProviderOptions
	static ProviderOptions_ClientSideProvider := 0x1
	static ProviderOptions_ServerSideProvider := 0x2
	static ProviderOptions_NonClientAreaProvider := 0x4
	static ProviderOptions_OverrideProvider := 0x8
	static ProviderOptions_ProviderOwnsSetFocus := 0x10
	static ProviderOptions_UseComThreading := 0x20
	static ProviderOptions_RefuseNonClientSupport := 0x40
	static ProviderOptions_HasNativeIAccessible := 0x80
	static ProviderOptions_UseClientCoordinates := 0x100

	ProviderOptions(Value="") {
		static v1:={0x1:"ClientSideProvider", 0x2:"ServerSideProvider", 0x4:"NonClientAreaProvider", 0x8:"OverrideProvider", 0x10:"ProviderOwnsSetFocus", 0x20:"UseComThreading", 0x40:"RefuseNonClientSupport", 0x80:"HasNativeIAccessible", 0x100:"UseClientCoordinates"}
		if Value is not integer
			return this["ProviderOptions_" Value]
		return (Value=="")?v1:v1[Value]
	}	

	; enum RowOrColumnMajor Contains values that specify whether data in a table should be read primarily by row or by column.	
	static RowOrColumnMajor_RowMajor := 0x0
	static RowOrColumnMajor_ColumnMajor := 0x1
	static RowOrColumnMajor_Indeterminate := 0x2

	RowOrColumnMajor(Value="") {
		static v1:={0x0:"RowMajor", 0x1:"ColumnMajor", 0x2:"Indeterminate"}
		if Value is not integer
			return this["RowOrColumnMajor_" Value]
		return (Value=="")?v1:v1[Value]
	}	
	
	; enum SayAsInterpretAs Defines the values that indicate how a text-to-speech engine should interpret specific data.
	static SayAsInterpretAs_None := 0
	static SayAsInterpretAs_Spell := 1
	static SayAsInterpretAs_Cardinal := 2
	static SayAsInterpretAs_Ordinal := 3
	static SayAsInterpretAs_Number := 4
	static SayAsInterpretAs_Date := 5
	static SayAsInterpretAs_Time := 6
	static SayAsInterpretAs_Telephone := 7
	static SayAsInterpretAs_Currency := 8
	static SayAsInterpretAs_Net := 9
	static SayAsInterpretAs_Url := 10
	static SayAsInterpretAs_Address := 11
	static SayAsInterpretAs_Alphanumeric := 12
	static SayAsInterpretAs_Name := 13
	static SayAsInterpretAs_Media := 14
	static SayAsInterpretAs_Date_MonthDayYear := 15
	static SayAsInterpretAs_Date_DayMonthYear := 16
	static SayAsInterpretAs_Date_YearMonthDay := 17
	static SayAsInterpretAs_Date_YearMonth := 18
	static SayAsInterpretAs_Date_MonthYear := 19
	static SayAsInterpretAs_Date_DayMonth := 20
	static SayAsInterpretAs_Date_MonthDay := 21
	static SayAsInterpretAs_Date_Year := 22
	static SayAsInterpretAs_Time_HoursMinutesSeconds12 := 23
	static SayAsInterpretAs_Time_HoursMinutes12 := 24
	static SayAsInterpretAs_Time_HoursMinutesSeconds24 := 25
	static SayAsInterpretAs_Time_HoursMinutes24 := 26
	
	SayAsInterpretAs(Value="") {
		static v1:={0:"None", 1:"Spell", 2:"Cardinal", 3:"Ordinal", 4:"Number", 5:"Date", 6:"Time", 7:"Telephone", 8:"Currency", 9:"Net", 10:"Url", 11:"Address", 13:"Name", 14:"Media", 15:"Date_MonthDayYear", 16:"Date_DayMonthYear", 17:"Date_YearMonthDay", 18:"Date_YearMonth", 19:"Date_MonthYear", 20:"Date_DayMonth", 21:"Date_MonthDay", 22:"Date_Year", 23:"Time_HoursMinutesSeconds12", 24:"Time_HoursMinutes12", 25:"Time_HoursMinutesSeconds24", 26:"Time_HoursMinutes24"}
		if Value is not integer
			return this["SayAsInterpretAs_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum ScrollAmount Contains values that specify the direction and distance to scroll.	
	static ScrollAmount_LargeDecrement := 0x0
	static ScrollAmount_SmallDecrement := 0x1
	static ScrollAmount_NoAmount := 0x2
	static ScrollAmount_LargeIncrement := 0x3
	static ScrollAmount_SmallIncrement := 0x4

	ScrollAmount(Value="") {
		static v1:={0x0:"LargeDecrement", 0x1:"SmallDecrement", 0x2:"NoAmount", 0x3:"LargeIncrement", 0x4:"SmallIncrement"}
		if Value is not integer
			return this["ScrollAmount_" Value]
		return (Value=="")?v1:v1[Value]
	}	
	
	; enum StructureChangeType Contains values that specify the type of change in the Microsoft UI Automation tree structure.	
	static StructureChangeType_ChildAdded := 0x0
	static StructureChangeType_ChildRemoved := 0x1
	static StructureChangeType_ChildrenInvalidated := 0x2
	static StructureChangeType_ChildrenBulkAdded := 0x3
	static StructureChangeType_ChildrenBulkRemoved := 0x4
	static StructureChangeType_ChildrenReordered := 0x5

	StructureChangeType(Value="") {
		static v1:={0x0:"ChildAdded", 0x1:"ChildRemoved", 0x2:"ChildrenInvalidated", 0x3:"ChildrenBulkAdded", 0x4:"ChildrenBulkRemoved", 0x5:"ChildrenReordered"}
		if Value is not integer
			return this["StructureChangeType_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum SupportedTextSelection Contains values that specify the supported text selection attribute.	
	static SupportedTextSelection_None := 0x0
	static SupportedTextSelection_Single := 0x1
	static SupportedTextSelection_Multiple := 0x2

	SupportedTextSelection(Value="") {
		static v1:={0x0:"None", 0x1:"Single", 0x2:"Multiple"}
		if Value is not integer
			return this["SupportedTextSelection_" Value]
		return (Value=="")?v1:v1[Value]
	}	

	; enum SynchronizedInputType Contains values that specify the type of synchronized input.	
	static SynchronizedInputType_KeyUp := 0x1
	static SynchronizedInputType_KeyDown := 0x2
	static SynchronizedInputType_LeftMouseUp := 0x4
	static SynchronizedInputType_LeftMouseDown := 0x8
	static SynchronizedInputType_RightMouseUp := 0x10
	static SynchronizedInputType_RightMouseDown := 0x20

	SynchronizedInputType(Value="") {
		static v1:={0x1:"KeyUp", 0x2:"KeyDown", 0x4:"LeftMouseUp", 0x8:"LeftMouseDown", 0x10:"RightMouseUp", 0x20:"RightMouseDown"}
		if Value is not integer
			return this["SynchronizedInputType_" Value]
		return (Value=="")?v1:v1[Value]
	}	

	; enum TextDecorationLineStyle Contains values that specify the OverlineStyle, StrikethroughStyle, and UnderlineStyle text attributes.
	static TextDecorationLineStyle_None := 0
	static TextDecorationLineStyle_Single := 1
	static TextDecorationLineStyle_WordsOnly := 2
	static TextDecorationLineStyle_Double := 3
	static TextDecorationLineStyle_Dot := 4
	static TextDecorationLineStyle_Dash := 5
	static TextDecorationLineStyle_DashDot := 6
	static TextDecorationLineStyle_DashDotDot := 7
	static TextDecorationLineStyle_Wavy := 8
	static TextDecorationLineStyle_ThickSingle := 9
	static TextDecorationLineStyle_DoubleWavy := 11
	static TextDecorationLineStyle_ThickWavy := 12
	static TextDecorationLineStyle_LongDash := 13
	static TextDecorationLineStyle_ThickDash := 14
	static TextDecorationLineStyle_ThickDashDot := 15
	static TextDecorationLineStyle_ThickDashDotDot := 16
	static TextDecorationLineStyle_ThickDot := 17
	static TextDecorationLineStyle_ThickLongDash := 18
	static TextDecorationLineStyle_Other := -1

	; enum TextEditChangeType Describes the text editing change being performed by controls when text-edit events are raised or handled.
	static TextEditChangeType_None := 0x0
	static TextEditChangeType_AutoCorrect := 0x1
	static TextEditChangeType_Composition := 0x2
	static TextEditChangeType_CompositionFinalized := 0x3
	static TextEditChangeType_AutoComplete := 0x4

	TextEditChangeType(Value="") {
		static v1:={0x0:"None", 0x1:"AutoCorrect", 0x2:"Composition", 0x3:"CompositionFinalized", 0x4:"AutoComplete"}
		if Value is not integer
			return this["TextEditChangeType_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum TextPatternRangeEndpoint Contains values that specify the endpoints of a text range.
	static TextPatternRangeEndpoint_Start := 0x0
	static TextPatternRangeEndpoint_End := 0x1

	TextPatternRangeEndpoint(Value="") {
		static v1:={0x0:"Start", 0x1:"End"}
		if Value is not integer
			return this["TextPatternRangeEndpoint_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum TextUnit Contains values that specify units of text for the purposes of navigation.	
	static TextUnit_Character := 0x0
	static TextUnit_Format := 0x1
	static TextUnit_Word := 0x2
	static TextUnit_Line := 0x3
	static TextUnit_Paragraph := 0x4
	static TextUnit_Page := 0x5
	static TextUnit_Document := 0x6

	TextUnit(Value="") {
		static v1:={0x0:"Character", 0x1:"Format", 0x2:"Word", 0x3:"Line", 0x4:"Paragraph", 0x5:"Page", 0x6:"Document"}
		if Value is not integer
			return this["TextUnit_" Value]
		return (Value=="")?v1:v1[Value]
	}	
	
	; enum ToggleState Contains values that specify the toggle state of a Microsoft UI Automation element that implements the Toggle control pattern.
	static ToggleState_Off := 0x0
	static ToggleState_On := 0x1
	static ToggleState_Indeterminate := 0x2

	ToggleState(Value="") {
		static v1:={0x0:"Off", 0x1:"On", 0x2:"Indeterminate"}
		if Value is not integer
			return this["ToggleState_" Value]
		return (Value=="")?v1:v1[Value]
	}

	;enum ZoomUnit Contains possible values for the IUIAutomationTransformPattern2::ZoomByUnit method, which zooms the viewport of a control by the specified unit.
	static ZoomUnit_NoAmount := 0x0
	static ZoomUnit_LargeDecrement := 0x1
	static ZoomUnit_SmallDecrement := 0x2
	static ZoomUnit_LargeIncrement := 0x3
	static ZoomUnit_SmallIncrement := 0x4

	ZoomUnit(Value="") {
		static v1:={0x0:"NoAmount", 0x1:"LargeDecrement", 0x2:"SmallDecrement", 0x3:"LargeIncrement", 0x4:"SmallIncrement"}
		if Value is not integer
			return this["ZoomUnit_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum WindowVisualState Contains values that specify the visual state of a window.
	static WindowVisualState_Normal := 0x0
	static WindowVisualState_Maximized := 0x1
	static WindowVisualState_Minimized := 0x2

	WindowVisualState(Value="") {
		static v1:={0x0:"Normal", 0x1:"Maximized", 0x2:"Minimized"}
		if Value is not integer
			return this["WindowVisualState_" Value]
		return (Value=="")?v1:v1[Value]
	}

	; enum WindowInteractionState Contains values that specify the current state of the window for purposes of user interaction.
	static WindowInteractionState_Running := 0x0
	static WindowInteractionState_Closing := 0x1
	static WindowInteractionState_ReadyForUserInteraction := 0x2
	static WindowInteractionState_BlockedByModalWindow := 0x3
	static WindowInteractionState_NotResponding := 0x4

	WindowInteractionState(Value="") {
		static v1:={0x0:"Running", 0x1:"Closing", 0x2:"ReadyForUserInteraction", 0x3:"BlockedByModalWindow", 0x4:"NotResponding"}
		if Value is not integer
			return this["WindowInteractionState_" Value]
		return (Value=="")?v1:v1[Value]
	}
}

class IAccessible_Enum { ; from UIA2.ahk (https://github.com/neptercn/UIAutomation/blob/master/UIA2.ahk)
	; enum SELFLAG specify how an accessible object becomes selected or takes the focus. http://msdn.microsoft.com/en-us/library/dd373634(v=vs.85).aspx
	static SELFLAG_NONE := 0
	static SELFLAG_TAKEFOCUS := 0x1
	static SELFLAG_TAKESELECTION := 0x2
	static SELFLAG_EXTENDSELECTION := 0x4
	static SELFLAG_ADDSELECTION := 0x8
	static SELFLAG_REMOVESELECTION := 0x10
	static SELFLAG_VALID := 0x1f
	
	; enum Roles describe the roles of various UI objects in an application. http://msdn.microsoft.com/en-us/library/dd373608%28v=vs.85%29.aspx
	static ROLE_SYSTEM_TITLEBAR := 0x1
	static ROLE_SYSTEM_MENUBAR := 0x2
	static ROLE_SYSTEM_SCROLLBAR := 0x3
	static ROLE_SYSTEM_GRIP := 0x4
	static ROLE_SYSTEM_SOUND := 0x5
	static ROLE_SYSTEM_CURSOR := 0x6
	static ROLE_SYSTEM_CARET := 0x7
	static ROLE_SYSTEM_ALERT := 0x8
	static ROLE_SYSTEM_WINDOW := 0x9
	static ROLE_SYSTEM_CLIENT := 0xa
	static ROLE_SYSTEM_MENUPOPUP := 0xb
	static ROLE_SYSTEM_MENUITEM := 0xc
	static ROLE_SYSTEM_TOOLTIP := 0xd
	static ROLE_SYSTEM_APPLICATION := 0xe
	static ROLE_SYSTEM_DOCUMENT := 0xf
	static ROLE_SYSTEM_PANE := 0x10
	static ROLE_SYSTEM_CHART := 0x11
	static ROLE_SYSTEM_DIALOG := 0x12
	static ROLE_SYSTEM_BORDER := 0x13
	static ROLE_SYSTEM_GROUPING := 0x14
	static ROLE_SYSTEM_SEPARATOR := 0x15
	static ROLE_SYSTEM_TOOLBAR := 0x16
	static ROLE_SYSTEM_STATUSBAR := 0x17
	static ROLE_SYSTEM_TABLE := 0x18
	static ROLE_SYSTEM_COLUMNHEADER := 0x19
	static ROLE_SYSTEM_ROWHEADER := 0x1a
	static ROLE_SYSTEM_COLUMN := 0x1b
	static ROLE_SYSTEM_ROW := 0x1c
	static ROLE_SYSTEM_CELL := 0x1d
	static ROLE_SYSTEM_LINK := 0x1e
	static ROLE_SYSTEM_HELPBALLOON := 0x1f
	static ROLE_SYSTEM_CHARACTER := 0x20
	static ROLE_SYSTEM_LIST := 0x21
	static ROLE_SYSTEM_LISTITEM := 0x22
	static ROLE_SYSTEM_OUTLINE := 0x23
	static ROLE_SYSTEM_OUTLINEITEM := 0x24
	static ROLE_SYSTEM_PAGETAB := 0x25
	static ROLE_SYSTEM_PROPERTYPAGE := 0x26
	static ROLE_SYSTEM_INDICATOR := 0x27
	static ROLE_SYSTEM_GRAPHIC := 0x28
	static ROLE_SYSTEM_STATICTEXT := 0x29
	static ROLE_SYSTEM_TEXT := 0x2a
	static ROLE_SYSTEM_PUSHBUTTON := 0x2b
	static ROLE_SYSTEM_CHECKBUTTON := 0x2c
	static ROLE_SYSTEM_RADIOBUTTON := 0x2d
	static ROLE_SYSTEM_COMBOBOX := 0x2e
	static ROLE_SYSTEM_DROPLIST := 0x2f
	static ROLE_SYSTEM_PROGRESSBAR := 0x30
	static ROLE_SYSTEM_DIAL := 0x31
	static ROLE_SYSTEM_HOTKEYFIELD := 0x32
	static ROLE_SYSTEM_SLIDER := 0x33
	static ROLE_SYSTEM_SPINBUTTON := 0x34
	static ROLE_SYSTEM_DIAGRAM := 0x35
	static ROLE_SYSTEM_ANIMATION := 0x36
	static ROLE_SYSTEM_EQUATION := 0x37
	static ROLE_SYSTEM_BUTTONDROPDOWN := 0x38
	static ROLE_SYSTEM_BUTTONMENU := 0x39
	static ROLE_SYSTEM_BUTTONDROPDOWNGRID := 0x3a
	static ROLE_SYSTEM_WHITESPACE := 0x3b
	static ROLE_SYSTEM_PAGETABLIST := 0x3c
	static ROLE_SYSTEM_CLOCK := 0x3d
	static ROLE_SYSTEM_SPLITBUTTON := 0x3e
	static ROLE_SYSTEM_IPADDRESS := 0x3f
	static ROLE_SYSTEM_OUTLINEBUTTON := 0x40
	
	; enum State describe the state of objects in an application UI. http://msdn.microsoft.com/en-us/library/dd373609%28v=vs.85%29.aspx
	static STATE_SYSTEM_NORMAL := 0
	static STATE_SYSTEM_UNAVAILABLE := 0x1
	static STATE_SYSTEM_SELECTED := 0x2
	static STATE_SYSTEM_FOCUSED := 0x4
	static STATE_SYSTEM_PRESSED := 0x8
	static STATE_SYSTEM_CHECKED := 0x10
	static STATE_SYSTEM_MIXED := 0x20
	static STATE_SYSTEM_INDETERMINATE := 0x20
	static STATE_SYSTEM_READONLY := 0x40
	static STATE_SYSTEM_HOTTRACKED := 0x80
	static STATE_SYSTEM_DEFAULT := 0x100
	static STATE_SYSTEM_EXPANDED := 0x200
	static STATE_SYSTEM_COLLAPSED := 0x400
	static STATE_SYSTEM_BUSY := 0x800
	static STATE_SYSTEM_FLOATING := 0x1000
	static STATE_SYSTEM_MARQUEED := 0x2000
	static STATE_SYSTEM_ANIMATED := 0x4000
	static STATE_SYSTEM_INVISIBLE := 0x8000
	static STATE_SYSTEM_OFFSCREEN := 0x10000
	static STATE_SYSTEM_SIZEABLE := 0x20000
	static STATE_SYSTEM_MOVEABLE := 0x40000
	static STATE_SYSTEM_SELFVOICING := 0x80000
	static STATE_SYSTEM_FOCUSABLE := 0x100000
	static STATE_SYSTEM_SELECTABLE := 0x200000
	static STATE_SYSTEM_LINKED := 0x400000
	static STATE_SYSTEM_TRAVERSED := 0x800000
	static STATE_SYSTEM_MULTISELECTABLE := 0x1000000
	static STATE_SYSTEM_EXTSELECTABLE := 0x2000000
	static STATE_SYSTEM_ALERT_LOW := 0x4000000
	static STATE_SYSTEM_ALERT_MEDIUM := 0x8000000
	static STATE_SYSTEM_ALERT_HIGH := 0x10000000
	static STATE_SYSTEM_PROTECTED := 0x20000000
	static STATE_SYSTEM_VALID := 0x7fffffff
	static STATE_SYSTEM_HASPOPUP := 0x40000000
	
	; enum Event describes the events that are generated by the operating system and by server applications. http://msdn.microsoft.com/en-us/library/dd318066%28v=vs.85%29.aspx
	static EVENT_MIN := 0x00000001
	static EVENT_MAX := 0x7FFFFFFF
	static EVENT_SYSTEM_SOUND := 0x0001
	static EVENT_SYSTEM_ALERT := 0x0002
	static EVENT_SYSTEM_FOREGROUND := 0x0003
	static EVENT_SYSTEM_MENUSTART := 0x0004
	static EVENT_SYSTEM_MENUEND := 0x0005
	static EVENT_SYSTEM_MENUPOPUPSTART := 0x0006
	static EVENT_SYSTEM_MENUPOPUPEND := 0x0007
	static EVENT_SYSTEM_CAPTURESTART := 0x0008
	static EVENT_SYSTEM_CAPTUREEND := 0x0009
	static EVENT_SYSTEM_MOVESIZESTART := 0x000A
	static EVENT_SYSTEM_MOVESIZEEND := 0x000B
	static EVENT_SYSTEM_CONTEXTHELPSTART := 0x000C
	static EVENT_SYSTEM_CONTEXTHELPEND := 0x000D
	static EVENT_SYSTEM_DRAGDROPSTART := 0x000E
	static EVENT_SYSTEM_DRAGDROPEND := 0x000F
	static EVENT_SYSTEM_DIALOGSTART := 0x0010
	static EVENT_SYSTEM_DIALOGEND := 0x0011
	static EVENT_SYSTEM_SCROLLINGSTART := 0x0012
	static EVENT_SYSTEM_SCROLLINGEND := 0x0013
	static EVENT_SYSTEM_SWITCHSTART := 0x0014
	static EVENT_SYSTEM_SWITCHEND := 0x0015
	static EVENT_SYSTEM_MINIMIZESTART := 0x0016
	static EVENT_SYSTEM_MINIMIZEEND := 0x0017
	static EVENT_SYSTEM_DESKTOPSWITCH := 0x0020
	static EVENT_SYSTEM_END := 0x00FF
	static EVENT_OEM_DEFINED_START := 0x0101
	static EVENT_OEM_DEFINED_END := 0x01FF
	static EVENT_UIA_EVENTID_START := 0x4E00
	static EVENT_UIA_EVENTID_END := 0x4EFF
	static EVENT_UIA_PROPID_START := 0x7500
	static EVENT_UIA_PROPID_END := 0x75FF
	static EVENT_CONSOLE_CARET := 0x4001
	static EVENT_CONSOLE_UPDATE_REGION := 0x4002
	static EVENT_CONSOLE_UPDATE_SIMPLE := 0x4003
	static EVENT_CONSOLE_UPDATE_SCROLL := 0x4004
	static EVENT_CONSOLE_LAYOUT := 0x4005
	static EVENT_CONSOLE_START_APPLICATION := 0x4006
	static EVENT_CONSOLE_END_APPLICATION := 0x4007
	static EVENT_CONSOLE_END := 0x40FF
	static EVENT_OBJECT_CREATE := 0x8000
	static EVENT_OBJECT_DESTROY := 0x8001
	static EVENT_OBJECT_SHOW := 0x8002
	static EVENT_OBJECT_HIDE := 0x8003
	static EVENT_OBJECT_REORDER := 0x8004
	static EVENT_OBJECT_FOCUS := 0x8005
	static EVENT_OBJECT_SELECTION := 0x8006
	static EVENT_OBJECT_SELECTIONADD := 0x8007
	static EVENT_OBJECT_SELECTIONREMOVE := 0x8008
	static EVENT_OBJECT_SELECTIONWITHIN := 0x8009
	static EVENT_OBJECT_STATECHANGE := 0x800A
	static EVENT_OBJECT_LOCATIONCHANGE := 0x800B
	static EVENT_OBJECT_NAMECHANGE := 0x800C
	static EVENT_OBJECT_DESCRIPTIONCHANGE := 0x800D
	static EVENT_OBJECT_VALUECHANGE := 0x800E
	static EVENT_OBJECT_PARENTCHANGE := 0x800F
	static EVENT_OBJECT_HELPCHANGE := 0x8010
	static EVENT_OBJECT_DEFACTIONCHANGE := 0x8011
	static EVENT_OBJECT_ACCELERATORCHANGE := 0x8012
	static EVENT_OBJECT_INVOKED := 0x8013
	static EVENT_OBJECT_TEXTSELECTIONCHANGED := 0x8014
	static EVENT_OBJECT_CONTENTSCROLLED := 0x8015
	static EVENT_SYSTEM_ARRANGMENTPREVIEW := 0x8016
	static EVENT_OBJECT_END := 0x80FF
	static EVENT_AIA_START := 0xA000
	static EVENT_AIA_END := 0xAFFF
	
	; enum Navigation indicate the spatial (up, down, left, and right) or logical (first child, last, next, and previous) direction observed when clients use IAccessible::accNavigate to navigate from one user interface element to another within the same container. http://msdn.microsoft.com/en-us/library/dd373600%28v=vs.85%29.aspx
	static NAVDIR_MIN := 0
	static NAVDIR_UP := 1
	static NAVDIR_DOWN := 2
	static NAVDIR_LEFT := 3
	static NAVDIR_RIGHT := 4
	static NAVDIR_NEXT := 5
	static NAVDIR_PREVIOUS := 6
	static NAVDIR_FIRSTCHILD := 7
	static NAVDIR_LASTCHILD := 8
	static NAVDIR_MAX := 9
	
	; enum Object Identifiers determine the object to which a WM_GETOBJECT message request refers. http://msdn.microsoft.com/en-us/library/dd373606%28v=vs.85%29.aspx
	static OBJID_WINDOW := 0
	static OBJID_SYSMENU := -1
	static OBJID_TITLEBAR := -2
	static OBJID_MENU := -3
	static OBJID_CLIENT := -4
	static OBJID_VSCROLL := -5
	static OBJID_HSCROLL := -6
	static OBJID_SIZEGRIP := -7
	static OBJID_CARET := -8
	static OBJID_CURSOR := -9
	static OBJID_ALERT := -10
	static OBJID_SOUND := -11
	static OBJID_QUERYCLASSNAMEIDX := -12
	static OBJID_NATIVEOM := -16
}

