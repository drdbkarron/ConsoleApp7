using System;
using System . Collections . Generic;
using System . Linq;
using System . Text;
using System . Threading . Tasks;
using Microsoft.Office.Core;
using System . Linq . Expressions;
using Microsoft . Office . Interop . Outlook;
using Microsoft . Office . Interop . Excel;
using System .Diagnostics;
using System . Runtime . InteropServices;
using System . Runtime . InteropServices . ComTypes;
using System . Diagnostics . Contracts;
using System . Windows . Controls . Ribbon;
using Microsoft . Office . Tools . Excel;
using Microsoft . Office . Tools . Ribbon;
using System . ComponentModel . Design;
using System . Reflection;
using Microsoft . Office . Interop . Access;

namespace ConsoleApp7
{
	class Program
	{
		static void Main ( string [ ] args )
		{
			ExcelChops ( );
			//OutlookChops ( );
		}

		private static void ExcelChops ( )
		{
			Process [ ] Running = Process . GetProcessesByName ( "Excel" );
			if ( Running . Count()==0 )
			{
				return;
			}

			Microsoft . Office . Interop . Excel . Application ExcelApplication = ( Microsoft . Office . Interop . Excel . Application ) Marshal . GetActiveObject ( "Excel.Application" );
			if ( ExcelApplication == null )
			{
				return;
			}

			string ActiveExcelApplicationCaption = ExcelApplication . Caption;
			Windows ExcelWindows = ExcelApplication . Windows;
			int ExcelWindowCount = ExcelWindows . Count;
			XlWindowState WindowState = ExcelApplication . WindowState;
			Window ExcelWindow = ExcelApplication . Windows [ 1 ];
			String ExcelWindoWindowCaption = ExcelWindow . Caption;

			System . Diagnostics . Debug . WriteLine ( String . Format ( "\nExcel Application Caption {0} " , ActiveExcelApplicationCaption ) );
			System . Diagnostics . Debug . WriteLine ( String . Format ( "\nExcel Window Caption {0} " , ExcelWindoWindowCaption ) );
			System . Diagnostics . Debug . WriteLine ( String . Format ( "Excel Window Count {0} " , ExcelWindowCount ) );
			System . Diagnostics . Debug . WriteLine ( String . Format ( "Excel Window State {0} " , WindowState ) );
			Microsoft.Office.Interop.Excel.Panes panes = ExcelWindow . Panes;
			IteratePanes ( panes );
			;

			Microsoft.Office.Interop.Excel.MenuBar aMB = ExcelApplication . ActiveMenuBar;
			IterateMenus ( aMB , 0 );
		System . Diagnostics . Debug . WriteLine ( String . Format ( "{0} {1} " , "Completed" , ( ( ( System . Environment . StackTrace ) . Split ( '\n' ) ) [ 2 ] . Trim ( ) ) ) );

		}

		private static void IteratePanes ( Microsoft . Office . Interop . Excel . Panes panes )
		{
		
			var n = panes . Count;
			for ( int i = 1 ; i <= n ; i++ )
			{

				System . Diagnostics . Debug . WriteLine ( String . Format ( "{0} {1} " , "panes" , ( ( ( System . Environment . StackTrace ) . Split ( '\n' ) ) [ 2 ] . Trim ( ) ) ) );
			}
		}

		private static void IterateMenus ( MenuBar aMB , int v )
		{

			string caption = aMB . Caption;
			int ndx = aMB . Index;
			dynamic parent = aMB . Parent;
			Menus menus = aMB . Menus;
			int menusCount = aMB . Menus . Count;

			for ( int i = 1 ; i <= menusCount ; i++ )
			{
				Menu a = menus [ i ];
				int b = a . Index;
				string c = a . Caption;
				System . Diagnostics . Debug . WriteLine ( String . Format ( "{0} {1} " , b , c ) );
				IterateMenus ( a , v + 1 );
			}

		}

		private static void IterateMenus ( Menu A , int v )
		{
			string caption = A . Caption;
			int ndx = A . Index;
			MenuItems items = A . MenuItems;
			int itemsCount = items . Count;

			for ( int i = 1 ; i <= itemsCount ; i++ )
			{
				dynamic a = items [ i ];
				Type t = a.GetType ( );

				object o = a as object;
				Type to = o . GetType ( );
				String oo = to . ToString ( );
				var occ = to . Name;
				var ooc = to . TypeHandle;

				System . Diagnostics . Debug . WriteLine ( String . Format ( "menu item {0} of {1} {2} {3} " , i , itemsCount, occ, caption) );
			}
		}

			private static void InstallRibbonTab ( Window ewI , string captionString , Microsoft . Office . Interop . Excel . Application eA )
		{

			dynamic a = eA . ActiveSheet;
			Pane aa = ewI . ActivePane;
			Microsoft . Office . Interop . Excel . Workbook b = eA . ActiveWorkbook;
			MenuBar c = eA . ActiveMenuBar;
			String Caption = ewI . Caption;
			Debug . WriteLine ( String . Format ( "Window #		{0} " ,		ewI . WindowNumber  ) );
			Debug . WriteLine ( String . Format ( "Caption		{0} " , Caption  ) );
			Debug . WriteLine ( String . Format ( "Window State	{0} " , ewI . WindowState ) );
			Debug . WriteLine ( String . Format ( "Names.Count  {0} " , eA.Names.Count) );
			Ribbon uuu = eA . ActiveMenuBar as Ribbon;
			foreach ( object n in eA . Names )
			{
				String nameString = n . ToString ( );
				Debug . WriteLine ( String . Format ( "Application.Name	{0} " , nameString ) );
			}
			Debug . WriteLine ( String . Format ( "pane index	{0} " , aa . Index ) );
			Debug . WriteLine ( String . Format ( "Menus Count	{0} " , c.Menus.Count ) );
			Menus ccc = c . Menus;

			foreach ( object ra in ccc)
			{
				RibbonApplicationMenuItem dd = ra as RibbonApplicationMenuItem;
				String ddd = dd . ToolTipDescription;

				System . Diagnostics . Debug . WriteLine ( String . Format ( "{0} {1} " , dd,  ( ( ( System . Environment . StackTrace ) . Split ( '\n' ) ) [ 2 ] . Trim ( ) ) ) );
			}
			var d = c . Menus;
			var e = d . Add ( "my Menu" );
			var f = e . MenuItems . Add ( "my Menu Items Added" );
			ewI . Activate ( );
			aa . Activate ( );

		}

		private static void OutlookChops ( )
		{
			Process [ ] Running = Process . GetProcessesByName ( "Outlook" );
			if ( Running . Count ( ) == 0 )
				return;

			Microsoft . Office . Interop . Outlook . Application OutlookApp = ( Microsoft . Office . Interop . Outlook . Application ) System . Runtime . InteropServices . Marshal . GetActiveObject ( "Outlook.Application" );

			if ( OutlookApp == null )

			{

				return;
			}

			System . Diagnostics . Debug . WriteLine ( string . Format ( "{0} {1} " , arg0: OutlookApp . Name , arg1: Environment . StackTrace . Split ( '\n' ) [ 2 ] . Trim ( ) ) );
			var ActiveE = OutlookApp . ActiveExplorer ( );
			//EnumerateFoldersInDefaultStore (OutlookApp);

		}

		static void EnumerateFoldersInDefaultStore ( Microsoft.Office.Interop.Outlook.Application OutlookApp)
		{
			NameSpace Sessions = OutlookApp . Session;
			Microsoft . Office . Interop . Outlook . Folder root = Sessions .
			DefaultStore . GetRootFolder ( ) as Microsoft . Office . Interop . Outlook . Folder;
			EnumerateFolders ( root );
		}

		// Uses recursion to enumerate Outlook sub-folders.
		static private void EnumerateFolders ( Microsoft . Office . Interop . Outlook . MAPIFolder folder )
		{
			Microsoft . Office . Interop . Outlook . Folders childFolders = folder . Folders;
			if ( childFolders . Count == 0 )
				{
				Debug . WriteLine ( String.Format("{0}:0",folder . FolderPath ));
				}
			else
				{
				
				foreach ( Microsoft . Office . Interop . Outlook . MAPIFolder childFolder in childFolders )
				{
					// Write the folder path.
					int items=childFolder.Items.Count;
					Debug . WriteLine ( String.Format("{0}:{1}", childFolder . FolderPath, items ) );
					childFolder . ShowItemCount= Microsoft . Office . Interop . Outlook . OlShowItemCount.olShowTotalItemCount;
					EnumerateFolders ( childFolder );
				}
			}
		}
	}
}
