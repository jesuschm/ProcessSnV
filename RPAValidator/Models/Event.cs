using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace RPAValidator.Model
{
    class Event
    {
        private String _Id;
        public enum EventType { Cursor, Keystrokes };
        private EventType _Event_type;
        private Point _Click_coord;
        private String _Event_info;
        private String _PicPath;

        public string Id { get => _Id; set => _Id = value; }
        public EventType Event_type { get => _Event_type; set => _Event_type = value; }
        public Point Click_coord { get => _Click_coord; set => _Click_coord = value; }
        public string Event_info { get => _Event_info; set => _Event_info = value; }
        public string PicPath { get => _PicPath; set => _PicPath = value; }

        public Event(List<String> values)
        {
            Id = values[0];

            Click_coord = new Point(Int32.Parse(values[1]), Int32.Parse(values[2]));

            Event_type = EventType.Cursor;
            if (values[3].Equals(EventType.Keystrokes.ToString("g")))
                Event_type = EventType.Keystrokes;

            Event_info = values[4];
            PicPath = values[5];
        }
        public bool IsCursor()
        {
            return Event_type.Equals(Event.EventType.Cursor);
        }
        public bool IsKeystroke()
        {
            return Event_type.Equals(Event.EventType.Keystrokes);
        }
    }
}
