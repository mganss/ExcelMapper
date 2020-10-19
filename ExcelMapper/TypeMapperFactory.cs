using System;
using System.Collections.Generic;
using System.Dynamic;

namespace Ganss.Excel
{
    /// <summary>
    /// A caching factory of <see cref="TypeMapper"/> objects.
    /// </summary>
    public class TypeMapperFactory : ITypeMapperFactory
    {
        Dictionary<Type, TypeMapper> TypeMappers { get; set; } = new Dictionary<Type, TypeMapper>();

        /// <summary>
        /// Creates a <see cref="TypeMapper"/> for the specified type.
        /// </summary>
        /// <param name="type">The type to create a <see cref="TypeMapper"/> object for.</param>
        /// <returns>A <see cref="TypeMapper"/> for the specified type.</returns>
        public TypeMapper Create(Type type)
        {
            if (!TypeMappers.TryGetValue(type, out TypeMapper typeMapper))
                typeMapper = TypeMappers[type] = TypeMapper.Create(type);

            return typeMapper;
        }

        /// <summary>
        /// Creates a <see cref="TypeMapper"/> for the specified object.
        /// </summary>
        /// <param name="o">The object to create a <see cref="TypeMapper"/> object for.</param>
        /// <returns>A <see cref="TypeMapper"/> for the specified object.</returns>
        public TypeMapper Create(object o)
        {
            if (o is ExpandoObject eo)
            {
                return TypeMapper.Create(eo);
            }
            else
            {
                return Create(o.GetType());
            }
        }
    }
}
