﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace FPBInteropConsole.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.6.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("\"no-reply@wufoo.com\"")]
        public string WufooSenderEmail {
            get {
                return ((string)(this["WufooSenderEmail"]));
            }
            set {
                this["WufooSenderEmail"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("secureorders@fergusonplarre.com.au")]
        public string MagentoSenderEmail {
            get {
                return ((string)(this["MagentoSenderEmail"]));
            }
            set {
                this["MagentoSenderEmail"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute(@"<?xml version=""1.0"" encoding=""utf-16""?>
<ArrayOfString xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
  <string>Custom Cakes</string>
  <string>Custom Cake Order</string>
  <string>Wufoo</string>
  <string>Decorating Room</string>
  <string>Vegan / Flourless Cake Order</string>
  <string>Cookie Cake Order Form</string>
</ArrayOfString>")]
        public global::System.Collections.Specialized.StringCollection WufooSenders {
            get {
                return ((global::System.Collections.Specialized.StringCollection)(this["WufooSenders"]));
            }
            set {
                this["WufooSenders"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("0")]
        public byte Setting {
            get {
                return ((byte)(this["Setting"]));
            }
            set {
                this["Setting"] = value;
            }
        }
    }
}
