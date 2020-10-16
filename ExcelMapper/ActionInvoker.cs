using System;

namespace Ganss.Excel
{
    /// <summary>
    /// Abstract class action invoker
    /// </summary>
    public class ActionInvoker
    {
        /// <summary>
        /// Invoke from an unspecified <paramref name="obj"/> type
        /// </summary>
        /// <param name="obj">mapping instance class</param>
        /// <param name="index">index in the collection</param>
        public virtual void Invoke(object obj, int index) =>
            throw new NotImplementedException();

        /// <summary>
        /// <see cref="ActionInvokerImpl{T}"/> factory
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="AfterMappingAction"></param>
        /// <returns></returns>
        public static ActionInvoker CreateInstance<T>(Action<T, int> AfterMappingAction)
        {
            // instanciate concrete generic invoker
            var invokerType = typeof(ActionInvokerImpl<>);
            Type[] tType = { typeof(T) };
            Type constructed = invokerType.MakeGenericType(tType);
            object invokerInstance = Activator.CreateInstance(constructed, AfterMappingAction);
            return (ActionInvoker)invokerInstance;
        }
    }

    /// <summary>
    /// Generic form <see cref="ActionInvoker"/> 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ActionInvokerImpl<T> : ActionInvoker
        where T : class
    {
        /// <summary>
        /// ref to the mapping action.
        /// </summary>
        internal Action<T, int> mappingAction;

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="mappingAction"></param>
        public ActionInvokerImpl(Action<T, int> mappingAction)
        {
            this.mappingAction = mappingAction;
        }

        /// <summary>
        /// Invoke Generic Action
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="index"></param>
        public override void Invoke(object obj, int index)
        {
            if (mappingAction is null || obj is null) return;
            mappingAction((obj as T), index);
        }
    }
}