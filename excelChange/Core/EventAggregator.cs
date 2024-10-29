using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelChange.Core
{
    public class EventAggregator
    {
        public event Action SpecialEventOccurred;

        public void PublishSpecialEvent()
        {
            SpecialEventOccurred?.Invoke();
        }
        public event Action TabsaveEventOccurred;
        public void TabsaveEvent()
        {
            TabsaveEventOccurred?.Invoke();
        }
    }
}
