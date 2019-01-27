using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelCake.NoIntrusive
{
    public class ExcelObject
    {
        private static Dictionary<Type, Dictionary<string, PropertyInfo>> _PropertyInfo = new Dictionary<Type, Dictionary<string, PropertyInfo>>();

        public Dictionary<string, Dictionary<string,object>> DataEntity { get; private set; }
        public Dictionary<string, Dictionary<string, List<object>>> DataList { get; private set; }

        public ExcelObject(object dataSource)
        {
            AnalysisDataSource(dataSource);
        }

        public ExcelObject(DataSet dataSource)
        {
            AnalysisDataSource(dataSource);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        private void AnalysisDataSource(object model)
        {
            DataEntity = new Dictionary<string, Dictionary<string, object>>();
            DataEntity.Add(string.Empty, new Dictionary<string, object>());
            DataList = new Dictionary<string, Dictionary<string, List<object>>>();

            if (model == null)
            {
                return;
            }

            Type modelType = model.GetType();
            var tps = GetProperty(modelType);

            foreach (var pn in tps.Keys)
            {
                var tp = tps[pn];
                var tpv = tp.GetValue(model, null);

                if (tpv == null)
                {
                    DataEntity[string.Empty].Add(pn, null); 
                    continue;
                }
                var tpt = tpv.GetType();

                if (tpt.Name.StartsWith("List") == true)
                {
                    DataList.Add(pn.ToUpper(), new Dictionary<string, List<object>>());
                    var tpvs = tpv as IEnumerable;
                    foreach (var tpvi in tpvs)
                    {
                        var tpp = GetProperty(tpvi.GetType());
                        foreach (var tppp in tpp.Values)
                        {
                            if (DataList[pn.ToUpper()].ContainsKey(tppp.Name.ToUpper()) == false) DataList[pn.ToUpper()].Add(tppp.Name.ToUpper(), new List<object>());
                            DataList[pn.ToUpper()][tppp.Name.ToUpper()].Add(tppp.GetValue(tpvi, null));
                        }
                    }
                }
                else if (tpt.FullName.StartsWith("System.") == true)
                {
                    DataEntity[string.Empty].Add(pn.ToUpper(), tpv);
                }
                else
                {
                    DataEntity.Add(pn.ToUpper(), new Dictionary<string, object>());
                    var tpp = GetProperty(tpt);
                    foreach (var tppp in tpp.Values)
                    {
                        DataEntity[pn.ToUpper()].Add(tppp.Name.ToUpper(), tppp.GetValue(tpv, null));
                    }
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataSource"></param>
        private void AnalysisDataSource(DataSet dataSource)
        {
            DataEntity = new Dictionary<string, Dictionary<string, object>>();
            DataEntity.Add(string.Empty, new Dictionary<string, object>());
            DataList = new Dictionary<string, Dictionary<string, List<object>>>();

            if (dataSource == null || dataSource.Tables == null)
            {
                return;
            }

            foreach (DataTable t in dataSource.Tables)
            {
                if (t.Rows.Count == 1)
                {
                    DataEntity.Add(t.TableName.ToUpper(), new Dictionary<string, object>());
                    for (int c = 0; c < t.Columns.Count; c++)
                    {
                        DataEntity[t.TableName.ToUpper()].Add(t.Columns[c].ColumnName.ToUpper(), t.Rows[0][c]);
                    }
                }
                DataList.Add(t.TableName.ToUpper(), new Dictionary<string, List<object>>());
                for (int c = 0; c < t.Columns.Count; c++)
                {
                    DataList[t.TableName.ToUpper()].Add(t.Columns[c].ColumnName.ToUpper(), new List<object>());
                }
                for (int r = 0; r < t.Rows.Count; r++)
                {
                    for (int c = 0; c < t.Columns.Count; c++)
                    {
                        DataList[t.TableName.ToUpper()][t.Columns[c].ColumnName.ToUpper()].Add(t.Rows[r][c]);
                    }
                }
            }
        }

        private static Dictionary<string, PropertyInfo> GetProperty(Type type)
        {
            if (_PropertyInfo.ContainsKey(type) == true)
            {
                return _PropertyInfo[type];
            }
            _PropertyInfo.Add(type, new Dictionary<string, PropertyInfo>());

            var tps = type.GetProperties();
            for (int i = 0; i < tps.Length; i++)
            {
                _PropertyInfo[type].Add(tps[i].Name, tps[i]);
            }
            return _PropertyInfo[type];
        }
    }
}
