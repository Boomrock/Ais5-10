using AngleSharp.Dom;
using AngleSharp;
using HTML_Parser.Models;
using System.Net.Http.Headers;


namespace HTML_Parser
{
    internal class Program
    {
        static void Main()
        {
            string url = "https://rasp.tpu.ru/";
            IConfiguration config = Configuration.Default.WithDefaultLoader();
            IBrowsingContext context = BrowsingContext.New(config);
            var task = Parse(url, context);
            task.Wait();
            using(var db = new LocalDBContext())
            {
     
                foreach (var group in db.Groups) {

                    Console.WriteLine("-----------------------------------------------------");
                    Console.Write($"School: {group.School}\n");
                    Console.Write($"Course: {group.Course}\n");
                    Console.Write($"Group: {group.Groupname}\n");
                    Console.Write($"Group link: {group.Link}\n");
                }
            }
            Console.ReadLine();
        }

        public static async Task<List<Group>> ParseSchool(IBrowsingContext context, string School, string Href)
        {
            var groups = new List<Group>();
            IDocument docSchool = await context.OpenAsync(Href);
            var aListCourseName = docSchool.QuerySelectorAll("ul.nav.nav-tabs.nav-tabs-bottom li a"); // name
            var aListCourseHref = docSchool.QuerySelectorAll("div.col-md-9 > ul.nav a").Select(elem => docSchool.Origin + elem.GetAttribute("href")); //href elem=> doc.Origin +...
            var CourseList = aListCourseName.Zip(aListCourseHref, (course, href) => new { Course = course.TextContent.Trim().Replace("\n", ""), Href = href });

            foreach (var a_course in CourseList)
            {
                IDocument docCourse = await context.OpenAsync(a_course.Href);
                var GroupHref = docCourse.QuerySelectorAll("div.tabs__content a").Select(elem => elem.GetAttribute("href")); //href
                var GroupName = docCourse.QuerySelectorAll("li.group-catalog__list-item a"); // name
                var GroupList = GroupName.Zip(GroupHref, (group, href) => new { Group = group.TextContent.Trim().Replace("\n", ""), Href = href });

                foreach (var a_group in GroupList)
                {
     
                    Group _group = new Group
                    {
                        School = School,
                        Course = a_course.Course,
                        Groupname = a_group.Group,
                        Link = a_group.Href
                    };
                    groups.Add(_group);
                }
            }
            return groups;
        }

        public static async Task Parse(string url, IBrowsingContext context)
        {
            IDocument docGlobal = await context.OpenAsync(url);

            var aListSchoolName = docGlobal.QuerySelectorAll("div.departments-list div"); // name
            var aListSchoolHref = docGlobal.QuerySelectorAll("div.departments-list a").Select(elem => elem.GetAttribute("href")); //href
            var SchoolList = aListSchoolName.Zip(aListSchoolHref, (school, href) => new { School = school.TextContent.Trim().Replace("\n", ""), Href = href });
            using (var dbCon = new LocalDBContext())
            {
                foreach (var a_school in SchoolList)
                {
                    List<Group> _groups = ParseSchool(context, a_school.School, a_school.Href).Result;
                    try
                    {
                        dbCon.Groups.AddRange(_groups);
                        dbCon.SaveChanges();
                    }
                    catch (Exception ex)
                    {
                        {
                            ex.ToString();
                            Console.WriteLine(ex.Message);
                            Console.Write(ex.ToString());
                        }
                    }

                }
            };
            Console.WriteLine("that all.");
        }
    }
}