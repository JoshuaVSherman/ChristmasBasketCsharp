﻿#pragma checksum "..\..\WindowSelectYear.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "F8B6621F5ACCFC149AE1341559CE6B9F"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.34014
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.Integration;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;


namespace ChristmasBasketsDashboard {
    
    
    /// <summary>
    /// WindowSelectYear
    /// </summary>
    public partial class WindowSelectYear : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 14 "..\..\WindowSelectYear.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListBox YearsListBox;
        
        #line default
        #line hidden
        
        
        #line 15 "..\..\WindowSelectYear.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button DeleteYear;
        
        #line default
        #line hidden
        
        
        #line 16 "..\..\WindowSelectYear.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button SelectYear;
        
        #line default
        #line hidden
        
        
        #line 17 "..\..\WindowSelectYear.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox YearToCreateTextBox;
        
        #line default
        #line hidden
        
        
        #line 18 "..\..\WindowSelectYear.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button CreateYear;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/Christmas Basket Dashboard;component/windowselectyear.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\WindowSelectYear.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.YearsListBox = ((System.Windows.Controls.ListBox)(target));
            
            #line 14 "..\..\WindowSelectYear.xaml"
            this.YearsListBox.SelectionChanged += new System.Windows.Controls.SelectionChangedEventHandler(this.YearsListBox_SelectionChanged);
            
            #line default
            #line hidden
            return;
            case 2:
            this.DeleteYear = ((System.Windows.Controls.Button)(target));
            
            #line 15 "..\..\WindowSelectYear.xaml"
            this.DeleteYear.Click += new System.Windows.RoutedEventHandler(this.DeleteYear_Click);
            
            #line default
            #line hidden
            return;
            case 3:
            this.SelectYear = ((System.Windows.Controls.Button)(target));
            
            #line 16 "..\..\WindowSelectYear.xaml"
            this.SelectYear.Click += new System.Windows.RoutedEventHandler(this.SelectYear_Click);
            
            #line default
            #line hidden
            return;
            case 4:
            this.YearToCreateTextBox = ((System.Windows.Controls.TextBox)(target));
            
            #line 17 "..\..\WindowSelectYear.xaml"
            this.YearToCreateTextBox.TextChanged += new System.Windows.Controls.TextChangedEventHandler(this.YearToCreateTextBox_TextChanged);
            
            #line default
            #line hidden
            return;
            case 5:
            this.CreateYear = ((System.Windows.Controls.Button)(target));
            
            #line 18 "..\..\WindowSelectYear.xaml"
            this.CreateYear.Click += new System.Windows.RoutedEventHandler(this.CreateYear_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

