﻿using System;
using System.Collections.Generic;
using System.Reflection;

namespace PPTXExp.phCommandFactory {
    public class PhCommandFactory {
        private Dictionary<string, string> dic = new Dictionary<string, string>();

        private PhCommandFactory() {
        
        }

        private static PhCommandFactory instance = null;
        public static PhCommandFactory GetInstance() {
            if (instance == null) {
                instance = new PhCommandFactory();
            }
            return instance;
        }

        public void CreateCommandInstance(string cls_name, params Object[] parameters) {
            Type t = Type.GetType(cls_name);
            ConstructorInfo[] ci = t.GetConstructors(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            if (ci.Length > 0) {
                ConstructorInfo ctor = ci[0];
                phCommand.PhCommand cmd = (phCommand.PhCommand)ctor.Invoke(null);
                cmd.Exec(parameters);
            }
        }

        public string GetHandledPPTX(string name) {
            try {
                return dic[name];
            } catch(Exception ex) {
                //Console.Write(ex.StackTrace);
                return null;
            }
        }
    }
}
