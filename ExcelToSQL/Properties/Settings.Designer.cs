//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан программой.
//     Исполняемая версия:4.0.30319.42000
//
//     Изменения в этом файле могут привести к неправильной работе и будут потеряны в случае
//     повторной генерации кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ExcelToSQL.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.8.1.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Praktikum")]
        public string DB_name {
            get {
                return ((string)(this["DB_name"]));
            }
            set {
                this["DB_name"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute(@"<?xml version=""1.0"" encoding=""utf-16""?>
<ArrayOfString xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
  <string>C:\Users\gaxy\Desktop\test</string>
  <string>C:\Users\gaxy\Desktop\test2</string>
</ArrayOfString>")]
        public global::System.Collections.Specialized.StringCollection Path_to_folder_library {
            get {
                return ((global::System.Collections.Specialized.StringCollection)(this["Path_to_folder_library"]));
            }
            set {
                this["Path_to_folder_library"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("C:\\Users\\gaxy\\Desktop\\1")]
        public string Copy_from {
            get {
                return ((string)(this["Copy_from"]));
            }
            set {
                this["Copy_from"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("C:\\Users\\gaxy\\Desktop\\test1")]
        public string Copy_to {
            get {
                return ((string)(this["Copy_to"]));
            }
            set {
                this["Copy_to"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool open {
            get {
                return ((bool)(this["open"]));
            }
            set {
                this["open"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Data Source=NOTEPC\\SQLEXPRESS;Initial Catalog=praktikum;Integrated Security=True")]
        public string SQL_string {
            get {
                return ((string)(this["SQL_string"]));
            }
            set {
                this["SQL_string"] = value;
            }
        }
    }
}
