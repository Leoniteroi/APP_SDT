Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.Runtime
Imports SDT.Civil.Commands

<Assembly: AssemblyTitle("SDT.Civil.Commands")>
<Assembly: AssemblyDescription("SDT toolkit for Civil 3D 2024")>
<Assembly: AssemblyCompany("SJTX Engenharia")>
<Assembly: AssemblyProduct("SDT – Civil 3D Toolkit")>
<Assembly: AssemblyCopyright("© 2025 SJTX Engenharia")>
<Assembly: AssemblyTrademark("")>

<Assembly: ComVisible(False)>
<Assembly: Guid("b5d0e2e0-efdf-4625-a9e6-5325cbe5e00d")>

<Assembly: AssemblyVersion("1.0.0.0")>
<Assembly: AssemblyFileVersion("1.0.0.0")>

' Registro de comandos (AutoCAD .NET)
<Assembly: CommandClass(GetType(MyCommands))>
<Assembly: CommandClass(GetType(Cmd_ListarBoundaries))>
<Assembly: CommandClass(GetType(SurfaceCommands))>
<Assembly: CommandClass(GetType(LabelAssembly))>

