using System;
using System.Drawing;

namespace KakaotalkBot
{
    internal static class ModerationParticipantSearch
    {
        internal static Rectangle Avatar(ParticipantMember member, Rectangle viewport)
        { var bounds = member.AvatarBounds; bounds.Offset(viewport.Left, viewport.Top - member.ScrollY); return bounds; }
        internal static int ScrollStep(Rectangle avatar, Rectangle viewport)
        {
            int distance = avatar.Top < viewport.Top ? viewport.Top - avatar.Top : avatar.Bottom - viewport.Bottom;
            return distance > viewport.Height ? 360 : 120;
        }
        internal static Rectangle Target(ParticipantMember member, long userId, string nickname, Rectangle viewport)
        {
            if (member == null || member.UserId != userId || member.Name != nickname || member.ScrollY < 0 || member.ViewportHeight != viewport.Height) return Rectangle.Empty;
            var avatar = Avatar(member, viewport); var label = member.NameBounds;
            label.Offset(viewport.Left, viewport.Top - member.ScrollY);
            if (avatar.Width < 12 || avatar.Width > 128 || avatar.Height < 12 || avatar.Height > 128 ||
                label.Width < 1 || label.Height < 1 || label.Height > 128 || label.Left < avatar.Right ||
                Math.Abs(label.Top + label.Height / 2 - (avatar.Top + avatar.Height / 2)) > avatar.Height / 2 ||
                !viewport.Contains(avatar) || label.Top < viewport.Top || label.Bottom > viewport.Bottom) return Rectangle.Empty;
            return avatar;
        }
    }
}
