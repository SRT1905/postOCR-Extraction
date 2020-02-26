namespace SmartOCR
{
    public interface IBuilder<T>
    {
        T Build();

        void Reset();
    }
}
