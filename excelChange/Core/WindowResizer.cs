using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;


namespace excelChange.Core
{
    public class WindowResizer
    {
        public enum ResizeDirection
        {
            None = 0,
            Left = 1,
            Right = 2,
            Top = 3,
            TopLeft = 4,
            TopRight = 5,
            Bottom = 6,
            BottomLeft = 7,
            BottomRight = 8,
            Drag = 10,
        }


        private const int WM_SYSCOMMAND = 0x112;
        /************************************************************************************
        * 함  수  명      : SendMessage
        * 내      용      : Windows API 호출을 통해 지정된 메시지를 윈도우에 전송합니다.
        * 설      명      : 이 함수는 특정 윈도우에 메시지를 전송하여 동작을 수행하도록 합니다.
        ************************************************************************************/

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        /************************************************************************************
        * 함  수  명      : GetDirection
        * 내      용      : 주어진 객체의 ResizeDirection 값을 가져옵니다.
        * 설      명      : 이 함수는 DependencyObject에서 ResizeDirection 속성을 가져옵니다.
        ************************************************************************************/

        public static ResizeDirection GetDirection(DependencyObject obj)
        {
            return (ResizeDirection)obj.GetValue(DirectionProperty);
        }
        /************************************************************************************
      * 함  수  명      : SetDirection
      * 내      용      : 주어진 객체에 ResizeDirection 값을 설정합니다.
      * 설      명      : 이 함수는 DependencyObject의 ResizeDirection 속성을 설정합니다.
      ************************************************************************************/
        public static void SetDirection(DependencyObject obj, ResizeDirection value)
        {
            obj.SetValue(DirectionProperty, value);
        }
        /************************************************************************************
        * 함  수  명      : DirectionProperty
        * 내      용      : ResizeDirection 속성을 등록합니다.
        * 설      명      : 이 속성은 ResizeDirection 값을 가진 DependencyProperty입니다.
        ************************************************************************************/

        public static readonly DependencyProperty DirectionProperty =
            DependencyProperty.RegisterAttached("Direction", typeof(ResizeDirection), typeof(WindowResizer),
            new UIPropertyMetadata(ResizeDirection.None, OnResizeDirectionChanged));

        /************************************************************************************
        * 함  수  명      : GetTargetWindow
        * 내      용      : 주어진 객체에 연결된 타겟 윈도우를 가져옵니다.
        * 설      명      : 이 함수는 DependencyObject에서 TargetWindow 속성을 가져옵니다.
        ************************************************************************************/

        private static Window GetTargetWindow(DependencyObject obj)
        {
            return (Window)obj.GetValue(TargetWindowProperty);
        }
        /************************************************************************************
        * 함  수  명      : TargetWindowProperty
        * 내      용      : TargetWindow 속성을 등록합니다.
        * 설      명      : 이 속성은 Window 값을 가진 DependencyProperty입니다.
        ************************************************************************************/

        private static readonly DependencyProperty TargetWindowProperty =
            DependencyProperty.RegisterAttached("TargetWindow", typeof(Window), typeof(ResizeDirection), new UIPropertyMetadata(null));
         /************************************************************************************
          * 함  수  명      : OnResizeDirectionChanged
          * 내      용      : ResizeDirection 속성이 변경될 때 호출됩니다.
          * 설      명      : 이 함수는 방향에 따라 커서를 설정하고 마우스 이벤트를 처리합니다.
          ************************************************************************************/

        protected static void OnResizeDirectionChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            FrameworkElement Target = sender as FrameworkElement;
            ResizeDirection Direction = (ResizeDirection)e.NewValue;

            switch (Direction)
            {
                case ResizeDirection.Left:
                case ResizeDirection.Right:
                    Target.Cursor = Cursors.SizeWE;
                    break;
                case ResizeDirection.Top:
                case ResizeDirection.Bottom:
                    Target.Cursor = Cursors.SizeNS;
                    break;
                case ResizeDirection.TopLeft:
                case ResizeDirection.BottomRight:
                    Target.Cursor = Cursors.SizeNWSE;
                    break;
                case ResizeDirection.TopRight:
                case ResizeDirection.BottomLeft:
                    Target.Cursor = Cursors.SizeNESW;
                    break;
                default:
                    Target.Cursor = Cursors.Arrow;
                    break;
            }

            Target.SetBinding(TargetWindowProperty, new Binding()
            { RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(Window), 1) });

            Target.MouseLeftButtonDown += (ss, ee) =>
            {
                if (Direction == ResizeDirection.Drag) DragWindow(Target);
                else ResizeWindow(Target, Direction);
                ee.Handled = true;
            };
        }
        /************************************************************************************
        * 함  수  명      : DragWindow
        * 내      용      : 윈도우를 드래그하여 이동합니다.
        * 설      명      : 이 함수는 주어진 FrameworkElement에 대해 DragMove를 호출하여 이동합니다.
        ************************************************************************************/
        private static void DragWindow(FrameworkElement Target)
        {
            Window Window = GetTargetWindow(Target);
            if (Window == null) return;
            Window.DragMove();
        }
        /************************************************************************************
         * 함  수  명      : ResizeWindow
         * 내      용      : 윈도우 크기를 조정합니다.
         * 설      명      : 이 함수는 ResizeDirection에 따라 윈도우의 크기를 조정합니다.
         ************************************************************************************/

        private static void ResizeWindow(FrameworkElement Target, ResizeDirection Direction)
        {
            if (Direction == ResizeDirection.None) return;

            Window Window = GetTargetWindow(Target);
            if (Window == null) return;

            Cursor CurrentCursor = Window.Cursor;
            Window.Cursor = Target.Cursor;

            HwndSource HwndSource = PresentationSource.FromVisual(Window) as HwndSource;
            SendMessage(HwndSource.Handle, WM_SYSCOMMAND, (IntPtr)(61440 + Direction), IntPtr.Zero);

            Window.Cursor = CurrentCursor;

        }




    }
}
