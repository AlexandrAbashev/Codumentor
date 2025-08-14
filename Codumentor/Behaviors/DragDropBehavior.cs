using System.Windows;
using System.Windows.Input;

namespace Codumentor
{
    public static class DragDropBehavior
    {
        public static readonly DependencyProperty DropCommandProperty =
            DependencyProperty.RegisterAttached(
                "DropCommand",
                typeof(ICommand),
                typeof(DragDropBehavior),
                new PropertyMetadata(null, OnDropCommandChanged));

        public static void SetDropCommand(DependencyObject element, ICommand value)
            => element.SetValue(DropCommandProperty, value);

        public static ICommand GetDropCommand(DependencyObject element)
            => (ICommand)element.GetValue(DropCommandProperty);

        private static void OnDropCommandChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is UIElement element)
            {
                if (e.NewValue != null)
                {
                    element.Drop += (sender, args) =>
                    {
                        var command = GetDropCommand(element);
                        if (command?.CanExecute(args) == true)
                            command.Execute(args);
                        args.Handled = true;
                    };
                }
            }
        }
    }
}