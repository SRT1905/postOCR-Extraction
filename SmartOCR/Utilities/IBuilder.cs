namespace SmartOCR
{
    interface IBuilder<T>
    {
        T Build();
        void Reset();
    }
}
