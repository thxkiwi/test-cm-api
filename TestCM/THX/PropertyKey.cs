using Windows.Win32.Devices.Properties;
using Windows.Win32.UI.Shell.PropertiesSystem;

namespace THX
{
    public class PropertyKey
    {
        private static DstT? GetField<SrcT, DstT>(SrcT? src, string fieldName, Func<object?, DstT?> f)
        {
            var fields = typeof(SrcT).GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .Where(f => f.Name.ToLower().Equals(fieldName.ToLower()));

            if (null != fields
                && fields.Count() > 0)
            {
                var field = fields.First();
                if (null == field)
                {
                    return f(null);
                }

                var val = field.GetValue(src);
                if (null == val)
                {
                    return f(null);
                }

                return f(val);
            }

            return f(null);
        }

        private static DstT? GetProperty<SrcT, DstT>(SrcT? src, string fieldNem, Func<object?, DstT?> f)
        {
            var props = typeof(SrcT).GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .Where(p => p.Name.ToLower().Equals(fieldNem.ToLower()));

            if (null != props
                && props.Count() > 0)
            {
                var prop = props.First();
                if (null == prop)
                {
                    return f(null);
                }

                var val = prop.GetValue(src);
                if (null == val)
                {
                    return f(null);
                }

                return f(val);
            }

            return f(null);
        }

        private static DstT? GetMethod<SrcT, DstT>(SrcT? src, string methodName, Func<object?, DstT?> f)
        {
            var methods = typeof(SrcT).GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .Where(m => m.Name.ToLower().Equals(methodName.ToLower()));

            if (null != methods
                               && methods.Count() > 0)
            {
                var method = methods.First();
                if (null == method)
                {
                    return f(null);
                }

                var val = method.Invoke(src, null);
                if (null == val)
                {
                    return f(null);
                }

                return f(val);
            }

            return f(null);
        }

        private static Guid? ToGuid(object? src)
        {
            if (src is Guid guid)
            {
                return guid;
            }

            if (null == src)
            {
                return null;
            }

            if (Guid.TryParse(src.ToString(), out guid))
            {
                return guid;
            }

            return null;
        }

        private static DstT? GetValue<SrcT, DstT>(SrcT t, string name, Func<object?, DstT?> f)
        {
            // try field
            DstT? val = GetField<SrcT?, DstT?>(t, name, f);
            if (null != val)
            {
                return val;
            }

            // try property
            val = GetProperty<SrcT?, DstT?>(t, name, f);
            if (null != val)
            {
                return val;
            }

            // try method
            val = GetMethod<SrcT?, DstT?>(t, name, f);
            if (null != val)
            {
                return val;
            }

            val = GetMethod<SrcT?, DstT?>(t, "get" + name, f);
            if (null != val)
            {
                return val;
            }

            return val;
        }

        public static PropertyKey From<ValueT>(ValueT key)
        {
            var fmtid = GetValue(key, "fmtid", new Func<object?, Guid?>(ToGuid));
            var pid = GetValue(key, "pid", new Func<object?, uint?>(o => (uint?)o));
            return new PropertyKey(fmtid, pid);
        }

        public Guid? fmtid { get; private set; }

        public uint? pid { get; private set; }

        public PropertyKey(Guid? fmtid, uint? pid)
        {
            this.fmtid = fmtid;
            this.pid = pid;
        }

        public override string? ToString()
        {
            return $"{fmtid?.ToString("B")},{pid}" ?? null;
        }

        public override bool Equals(object? obj)
        {
            if (obj is PropertyKey pk)
            {
                return fmtid == pk.fmtid && pid == pk.pid;
            }
            else if (obj is string s)
            {
                return s.Equals(ToString());
            }
            else if (obj is PROPERTYKEY pk1)
            {
                return fmtid == pk1.fmtid && pid == pk1.pid;
            }
            else if (obj is DEVPROPKEY dpk)
            {
                return fmtid == dpk.fmtid && pid == dpk.pid;
            }

            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(fmtid, pid);
        }
    }
}
