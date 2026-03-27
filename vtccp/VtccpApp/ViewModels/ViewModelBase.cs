namespace VtccpApp.ViewModels;

using System.ComponentModel;
using System.Runtime.CompilerServices;

/// <summary>
/// Base class that implements <see cref="INotifyPropertyChanged"/> for all view-models.
/// </summary>
public abstract class ViewModelBase : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>
    /// Sets <paramref name="field"/> to <paramref name="value"/> and raises
    /// <see cref="PropertyChanged"/> if the value changed.
    /// </summary>
    protected bool Set<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
