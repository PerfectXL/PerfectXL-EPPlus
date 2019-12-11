/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		10-SEP-2009
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: InternalsVisibleTo("PerfectXL.EPPlus.UnitTests")]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("9dd43b8d-c4fe-4a8b-ad6e-47ef83bbbb01")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Revision and Build Numbers 
// by using the '*' as shown below:
#if (!Core)
    //[assembly: AssemblyTitle("EPPlus")]
    //[assembly: AssemblyDescription("Allows Excel files(xlsx;xlsm) to be created on the server. See epplus.codeplex.com")]
    //[assembly: AssemblyConfiguration("")]
    //[assembly: AssemblyCompany("EPPlus")]
    //[assembly: AssemblyProduct("EPPlus")]
    //[assembly: AssemblyCopyright("Copyright 2009- ©Jan Källman. Parts of the Interface comes from the ExcelPackage-project")]
    //[assembly: AssemblyTrademark("The GNU Lesser General Public License (LGPL)")]
    //[assembly: AssemblyCulture("")]
    //[assembly: ComVisible(false)]

    //[assembly: AssemblyVersion("4.5.1")]
    //[assembly: AssemblyFileVersion("4.5.0.0")]
#endif
[assembly: AllowPartiallyTrustedCallers]
