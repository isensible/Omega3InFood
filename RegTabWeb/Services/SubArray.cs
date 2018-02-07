using System;
using System.Collections;
using System.Collections.Generic;

namespace RegTabWeb.Services
{
    public class SubArray<T> : IEnumerable<T>
    {
        private ArraySegment<T> _arraySegment;

        public SubArray(T[] array, int offset, int count)
        {
            _arraySegment = new ArraySegment<T>(array, offset, count);
        }
        
        public int Count => _arraySegment.Count;

        public int Offset => _arraySegment.Offset;

        public T this[int index] => _arraySegment.Array[_arraySegment.Offset + index];

        public T[] ToArray()
        {
            T[] temp = new T[_arraySegment.Count];
            Array.Copy(_arraySegment.Array, _arraySegment.Offset, temp, 0, _arraySegment.Count);
            return temp;
        }

        public IEnumerator<T> GetEnumerator()
        {
            for (int i = _arraySegment.Offset; i < _arraySegment.Offset + _arraySegment.Count; i++)
            {
                yield return _arraySegment.Array[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}