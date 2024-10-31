using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelChange.Core
{
    //이벤트의 구독(Subscribe)과 출판(Publish)의 매커니즘을 가지고 있는 구조
    public class EventAggregator
    {
        public event Action SpecialEventOccurred;

        /************************************************************************************
         * 함  수  명      : PublishSpecialEvent
         * 내      용      : 특수 이벤트를 발행합니다.
         * 설      명      : 이 함수는 SpecialEventOccurred 이벤트를 발생시킵니다.
        ************************************************************************************/

        public void PublishSpecialEvent()
        {
            SpecialEventOccurred?.Invoke();
        }

        // 탭 저장 이벤트가 발생했을 때 구독자에게 알리는 이벤트 (추후 사용될까봐)
        public event Action TabsaveEventOccurred;

        /************************************************************************************
        * 함  수  명      : TabsaveEvent
        * 내      용      : 탭 저장 이벤트를 발행합니다.
        * 설      명      : 이 함수는 TabsaveEventOccurred 이벤트를 발생시킵니다.
        ************************************************************************************/

        public void TabsaveEvent()
        {
            TabsaveEventOccurred?.Invoke();
        }
    }
}
