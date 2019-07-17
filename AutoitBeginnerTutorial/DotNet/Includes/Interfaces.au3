#include-once

Global Const $sCLSID_CorRuntimeHost = "{CB2F6723-AB3A-11D2-9C40-00C04FA30A3E}"
Global Const $tCLSID_CorRuntimeHost = CLSIDFromString( $sCLSID_CorRuntimeHost )
Global Const $sIID_ICorRuntimeHost = "{CB2F6722-AB3A-11D2-9C40-00C04FA30A3E}"
Global Const $tIID_ICorRuntimeHost = CLSIDFromString( $sIID_ICorRuntimeHost )
Global Const $sTag_ICorRuntimeHost = _
	"CreateLogicalThreadState hresult();" & _
	"DeleteLogicalThreadState hresult();" & _
	"SwitchInLogicalThreadState hresult();" & _
	"SwitchOutLogicalThreadState hresult();" & _
	"LocksHeldByLogicalThread hresult();" & _
	"MapFile hresult();" & _
	"GetConfiguration hresult();" & _
	"Start hresult();" & _
	"Stop hresult();" & _
	"CreateDomain hresult();" & _
	"GetDefaultDomain hresult(ptr*);" & _
	"EnumDomains hresult(ptr*);" & _
	"NextDomain hresult(ptr;ptr*);" & _
	"CloseEnum hresult();" & _
	"CreateDomainEx hresult();" & _
	"CreateDomainSetup hresult();" & _
	"CreateEvidence hresult();" & _
	"UnloadDomain hresult(ptr);" & _
	"CurrentDomain hresult();"

Global Const $sIID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
Global Const $sTag_IDispatch = _
	"GetTypeInfoCount hresult(dword*);" & _
	"GetTypeInfo hresult(dword;dword;ptr*);" & _
	"GetIDsOfNames hresult(ptr;ptr;dword;dword;ptr);" & _
	"Invoke hresult(dword;ptr;dword;word;ptr;ptr;ptr;ptr);"

; The interfaces _AppDomain, _Type and _Assembly below that starts with
; an underscore are the interfaces that can be used from unmanaged code.

Global Const $sIID__AppDomain = "{05F696DC-2B29-3663-AD8B-C4389CF2A713}"
Global Const $sTag__AppDomain = _
	$sTag_IDispatch & _
	"get_ToString hresult();" & _
	"Equals hresult();" & _
	"GetHashCode hresult();" & _
	"GetType hresult(ptr*);" & _
	"InitializeLifetimeService hresult();" & _
	"GetLifetimeService hresult();" & _
	"get_Evidence hresult();" & _
	"add_DomainUnload hresult();" & _
	"remove_DomainUnload hresult();" & _
	"add_AssemblyLoad hresult();" & _
	"remove_AssemblyLoad hresult();" & _
	"add_ProcessExit hresult();" & _
	"remove_ProcessExit hresult();" & _
	"add_TypeResolve hresult();" & _
	"remove_TypeResolve hresult();" & _
	"add_ResourceResolve hresult();" & _
	"remove_ResourceResolve hresult();" & _
	"add_AssemblyResolve hresult();" & _
	"remove_AssemblyResolve hresult();" & _
	"add_UnhandledException hresult();" & _
	"remove_UnhandledException hresult();" & _
	"DefineDynamicAssembly hresult();" & _
	"DefineDynamicAssembly_2 hresult();" & _
	"DefineDynamicAssembly_3 hresult();" & _
	"DefineDynamicAssembly_4 hresult();" & _
	"DefineDynamicAssembly_5 hresult();" & _
	"DefineDynamicAssembly_6 hresult();" & _
	"DefineDynamicAssembly_7 hresult();" & _
	"DefineDynamicAssembly_8 hresult();" & _
	"DefineDynamicAssembly_9 hresult();" & _
	"CreateInstance hresult(bstr;bstr;object*);" & _
	"CreateInstanceFrom hresult();" & _
	"CreateInstance_2 hresult();" & _
	"CreateInstanceFrom_2 hresult();" & _
	"CreateInstance_3 hresult(bstr;bstr;bool;int;ptr;ptr;ptr;ptr;ptr;ptr*);" & _
	"CreateInstanceFrom_3 hresult();" & _
	"Load hresult();" & _
	"Load_2 hresult();" & _
	"Load_3 hresult();" & _
	"Load_4 hresult();" & _
	"Load_5 hresult();" & _
	"Load_6 hresult();" & _
	"Load_7 hresult();" & _
	"ExecuteAssembly hresult();" & _
	"ExecuteAssembly_2 hresult();" & _
	"ExecuteAssembly_3 hresult();" & _
	"get_FriendlyName hresult(bstr*);" & _
	"get_BaseDirectory hresult(bstr*);" & _
	"get_RelativeSearchPath hresult();" & _
	"get_ShadowCopyFiles hresult();" & _
	"GetAssemblies hresult(ptr*);" & _
	"AppendPrivatePath hresult();" & _
	"ClearPrivatePath ) = 0; hresult();" & _
	"SetShadowCopyPath hresult();" & _
	"ClearShadowCopyPath ) = 0; hresult();" & _
	"SetCachePath hresult();" & _
	"SetData hresult();" & _
	"GetData hresult();" & _
	"SetAppDomainPolicy hresult();" & _
	"SetThreadPrincipal hresult();" & _
	"SetPrincipalPolicy hresult();" & _
	"DoCallBack hresult();" & _
	"get_DynamicDirectory hresult();"

Global Const $sIID__Type = "{BCA8B44D-AAD6-3A86-8AB7-03349F4F2DA2}"
Global Const $sTag__Type = _
	$sTag_IDispatch & _
	"get_ToString hresult(bstr*);" & _
	"Equals hresult(variant;short*);" & _
	"GetHashCode hresult(int*);" & _
	"GetType hresult(ptr);" & _
	"get_MemberType hresult(ptr);" & _
	"get_name hresult(bstr*);" & _
	"get_DeclaringType hresult(ptr);" & _
	"get_ReflectedType hresult(ptr);" & _
	"GetCustomAttributes hresult(ptr;short;ptr);" & _
	"GetCustomAttributes_2 hresult(short;ptr);" & _
	"IsDefined hresult(ptr;short;short*);" & _
	"get_Guid hresult(ptr);" & _
	"get_Module hresult(ptr);" & _
	"get_Assembly hresult(ptr*);" & _
	"get_TypeHandle hresult(ptr);" & _
	"get_FullName hresult(bstr*);" & _
	"get_Namespace hresult(bstr*);" & _
	"get_AssemblyQualifiedName hresult(bstr*);" & _
	"GetArrayRank hresult(int*);" & _
	"get_BaseType hresult(ptr);" & _
	"GetConstructors hresult(ptr;ptr);" & _
	"GetInterface hresult(bstr;short;ptr);" & _
	"GetInterfaces hresult(ptr);" & _
	"FindInterfaces hresult(ptr;variant;ptr);" & _
	"GetEvent hresult(bstr;ptr;ptr);" & _
	"GetEvents hresult(ptr);" & _
	"GetEvents_2 hresult(int;ptr);" & _
	"GetNestedTypes hresult(int;ptr);" & _
	"GetNestedType hresult(bstr;ptr;ptr);" & _
	"GetMember hresult(bstr;ptr;ptr;ptr);" & _
	"GetDefaultMembers hresult(ptr);" & _
	"FindMembers hresult(ptr;ptr;ptr;variant;ptr);" & _
	"GetElementType hresult(ptr);" & _
	"IsSubclassOf hresult(ptr;short*);" & _
	"IsInstanceOfType hresult(variant;short*);" & _
	"IsAssignableFrom hresult(ptr;short*);" & _
	"GetInterfaceMap hresult(ptr;ptr);" & _
	"GetMethod hresult(bstr;ptr;ptr;ptr;ptr;ptr);" & _
	"GetMethod_2 hresult(bstr;ptr;ptr);" & _
	"GetMethods hresult(int;ptr);" & _
	"GetField hresult(bstr;ptr;ptr);" & _
	"GetFields hresult(int;ptr);" & _
	"GetProperty hresult(bstr;ptr;ptr);" & _
	"GetProperty_2 hresult(bstr;ptr;ptr;ptr;ptr;ptr;ptr);" & _
	"GetProperties hresult(ptr;ptr);" & _
	"GetMember_2 hresult(bstr;ptr;ptr);" & _
	"GetMembers hresult(int;ptr*);" & _
	"InvokeMember hresult(bstr;ptr;ptr;variant;ptr;ptr;ptr;ptr;variant*);" & _
	"get_UnderlyingSystemType hresult(ptr);" & _
	"InvokeMember_2 hresult(bstr;int;ptr;variant;ptr;ptr;variant*);" & _
	"InvokeMember_3 hresult(bstr;int;ptr;variant;ptr;variant*);" & _
	"GetConstructor hresult(ptr;ptr;ptr;ptr;ptr;ptr);" & _
	"GetConstructor_2 hresult(ptr;ptr;ptr;ptr;ptr);" & _
	"GetConstructor_3 hresult(ptr;ptr);" & _
	"GetConstructors_2 hresult(ptr);" & _
	"get_TypeInitializer hresult(ptr);" & _
	"GetMethod_3 hresult(bstr;ptr;ptr;ptr;ptr;ptr;ptr);" & _
	"GetMethod_4 hresult(bstr;ptr;ptr;ptr);" & _
	"GetMethod_5 hresult(bstr;ptr;ptr);" & _
	"GetMethod_6 hresult(bstr;ptr);" & _
	"GetMethods_2 hresult(ptr);" & _
	"GetField_2 hresult(bstr;ptr);" & _
	"GetFields_2 hresult(ptr);" & _
	"GetInterface_2 hresult(bstr;ptr);" & _
	"GetEvent_2 hresult(bstr;ptr);" & _
	"GetProperty_3 hresult(bstr;ptr;ptr;ptr;ptr);" & _
	"GetProperty_4 hresult(bstr;ptr;ptr;ptr);" & _
	"GetProperty_5 hresult(bstr;ptr;ptr);" & _
	"GetProperty_6 hresult(bstr;ptr;ptr);" & _
	"GetProperty_7 hresult(bstr;ptr);" & _
	"GetProperties_2 hresult(ptr);" & _
	"GetNestedTypes_2 hresult(ptr);" & _
	"GetNestedType_2 hresult(bstr;ptr);" & _
	"GetMember_3 hresult(bstr;ptr);" & _
	"GetMembers_2 hresult(ptr);" & _
	"get_Attributes hresult(ptr);" & _
	"get_IsNotPublic hresult(short*);" & _
	"get_IsPublic hresult(short*);" & _
	"get_IsNestedPublic hresult(short*);" & _
	"get_IsNestedPrivate hresult(short*);" & _
	"get_IsNestedFamily hresult(short*);" & _
	"get_IsNestedAssembly hresult(short*);" & _
	"get_IsNestedFamANDAssem hresult(short*);" & _
	"get_IsNestedFamORAssem hresult(short*);" & _
	"get_IsAutoLayout hresult(short*);" & _
	"get_IsLayoutSequential hresult(short*);" & _
	"get_IsExplicitLayout hresult(short*);" & _
	"get_IsClass hresult(short*);" & _
	"get_IsInterface hresult(short*);" & _
	"get_IsValueType hresult(short*);" & _
	"get_IsAbstract hresult(short*);" & _
	"get_IsSealed hresult(short*);" & _
	"get_IsEnum hresult(short*);" & _
	"get_IsSpecialName hresult(short*);" & _
	"get_IsImport hresult(short*);" & _
	"get_IsSerializable hresult(short*);" & _
	"get_IsAnsiClass hresult(short*);" & _
	"get_IsUnicodeClass hresult(short*);" & _
	"get_IsAutoClass hresult(short*);" & _
	"get_IsArray hresult(short*);" & _
	"get_IsByRef hresult(short*);" & _
	"get_IsPointer hresult(short*);" & _
	"get_IsPrimitive hresult(short*);" & _
	"get_IsCOMObject hresult(short*);" & _
	"get_HasElementType hresult(short*);" & _
	"get_IsContextful hresult(short*);" & _
	"get_IsMarshalByRef hresult(short*);" & _
	"Equals_2 hresult(ptr;short*);"

; Binding flags for InvokeMember, InvokeMember_2
; and InvokeMember_3 methods of _Type interface.
Global Const $BindingFlags_Default = 0x0000
Global Const $BindingFlags_IgnoreCase = 0x0001
Global Const $BindingFlags_DeclaredOnly = 0x0002
Global Const $BindingFlags_Instance = 0x0004
Global Const $BindingFlags_Static = 0x0008
Global Const $BindingFlags_Public = 0x0010
Global Const $BindingFlags_NonPublic = 0x0020
Global Const $BindingFlags_FlattenHierarchy = 0x0040
Global Const $BindingFlags_InvokeMethod = 0x0100
Global Const $BindingFlags_CreateInstance = 0x0200
Global Const $BindingFlags_GetField = 0x0400
Global Const $BindingFlags_SetField = 0x0800
Global Const $BindingFlags_GetProperty = 0x1000
Global Const $BindingFlags_SetProperty = 0x2000
Global Const $BindingFlags_PutDispProperty = 0x4000
Global Const $BindingFlags_PutRefDispProperty = 0x8000
Global Const $BindingFlags_ExactBinding = 0x00010000
Global Const $BindingFlags_SuppressChangeType = 0x00020000
Global Const $BindingFlags_OptionalParamBinding = 0x00040000
Global Const $BindingFlags_IgnoreReturn = 0x01000000
Global Const $BindingFlags_DefaultValue = $BindingFlags_Static + $BindingFlags_Public + $BindingFlags_FlattenHierarchy + $BindingFlags_InvokeMethod

Global Const $sIID__Assembly = "{17156360-2F1A-384A-BC52-FDE93C215C5B}"
Global Const $sTag__Assembly = _
	$sTag_IDispatch & _
	"get_ToString hresult(bstr*);" & _
	"Equals hresult();" & _
	"GetHashCode hresult();" & _
	"GetType hresult(ptr*);" & _
	"get_CodeBase hresult();" & _
	"get_EscapedCodeBase hresult();" & _
	"GetName hresult();" & _
	"GetName_2 hresult();" & _
	"get_FullName hresult(bstr*);" & _
	"get_EntryPoint hresult();" & _
	"GetType_2 hresult(bstr;ptr*);" & _
	"GetType_3 hresult();" & _
	"GetExportedTypes hresult();" & _
	"GetTypes hresult(ptr*);" & _
	"GetManifestResourceStream hresult();" & _
	"GetManifestResourceStream_2 hresult();" & _
	"GetFile hresult();" & _
	"GetFiles hresult();" & _
	"GetFiles_2 hresult();" & _
	"GetManifestResourceNames hresult();" & _
	"GetManifestResourceInfo hresult();" & _
	"get_Location hresult(bstr*);" & _
	"get_Evidence hresult();" & _
	"GetCustomAttributes hresult();" & _
	"GetCustomAttributes_2 hresult();" & _
	"IsDefined hresult();" & _
	"GetObjectData hresult();" & _
	"add_ModuleResolve hresult();" & _
	"remove_ModuleResolve hresult();" & _
	"GetType_4 hresult();" & _
	"GetSatelliteAssembly hresult();" & _
	"GetSatelliteAssembly_2 hresult();" & _
	"LoadModule hresult();" & _
	"LoadModule_2 hresult();" & _
	"CreateInstance hresult(bstr;variant*);" & _
	"CreateInstance_2 hresult(bstr;bool;variant*);" & _
	"CreateInstance_3 hresult(bstr;bool;int;ptr;ptr;ptr;ptr;variant*);" & _
	"GetLoadedModules hresult();" & _
	"GetLoadedModules_2 hresult();" & _
	"GetModules hresult();" & _
	"GetModules_2 hresult();" & _
	"GetModule hresult();" & _
	"GetReferencedAssemblies hresult();" & _
	"get_GlobalAssemblyCache hresult(bool*);"

Func CLSIDFromString( $sGUID )
	Static $hOle32Dll = "ole32.dll", $tGUID = DllStructCreate( "ulong Data1;ushort Data2;ushort Data3;byte Data4[8]" ), $pGUID = DllStructGetPtr( $tGUID )
	DllCall( $hOle32Dll, "uint", "CLSIDFromString", "wstr", $sGUID, "ptr", $pGUID )
	Return $tGUID
EndFunc

Func GUIDFromStringEx( $sGUID, $tGUID )
	Static $hOle32Dll = "ole32.dll"
	DllCall( $hOle32Dll, "long", "CLSIDFromString", "wstr", $sGUID, "struct*", $tGUID )
EndFunc
