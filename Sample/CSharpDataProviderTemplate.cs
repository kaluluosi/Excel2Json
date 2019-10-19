using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LitJson;
using System.IO;
using UnityEngine;


    namespace Data{{


        public  class {ClassName}
        {{
            private static {ClassName} instance;
            public static {ClassName} Instance
            {{
                get
                {{
                    return instance == null ? instance = new {ClassName}() : instance;
                }}
            }}

            public Dictionary<string,DataNode> Data {{ get;private set; }}

            private {{ClassName}}()
            {{
                    TextAsset textAsset = CfgProvider.Get("{CfgName}");
                    Data = JsonMapper.ToObject<Dictionary<string,DataNode>>(textAsset.text);
            }}

            public DataNode this[int index]
            {{
                get
                {{
                    return Data[index.ToString()];
                }}
            }}


            public class DataNode
            {{
                {DataNodeFields}
            }}
        }}
    }}