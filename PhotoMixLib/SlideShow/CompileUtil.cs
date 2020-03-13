using System;
using System.Collections.Generic;
using System.Text;

namespace Msn.PhotoMix.SlideShow
{
    public interface ICompile
    {
        bool IsFixedCount { get; }      
        bool CompileAll { get; set; }
        int CompileMaxCount { get; set; }
        int CompileMinCount { get; set; }
        CompileOrder CompileOrder { get; set; }
        CompileInclude CompileInclude { get; set; }
        int CompileAge { get; set; }
    }

    public enum CompileOrder
    {
        Listed = 0,
        Sorted = 1,
        Random = 2,
        DailyInterval = 3
    }

    public enum CompileInclude
    {
        All = 0,
        Days = 1
    }

    public class CompileUtil
    {
        static public void PruneList(List<ListItem> inputList, bool all, int maxCount, int minCount, CompileOrder order, CompileInclude includeSince, int ageCutoff, DateTime currentDate, DateTime seedDate)
        {
            // Remove anything older than the age cuttoff
            if (includeSince == CompileInclude.Days && ageCutoff != 0)
            {
                int i = 0;
                while (i < inputList.Count)
                {
                    ListItem item = inputList[i];

                    DateTime pubDate = item.PubDate;
                    TimeSpan ts = currentDate.Subtract(pubDate);
                    if (ts.Days > ageCutoff)
                    {
                        inputList.Remove(item);
                    }
                    else
                    {
                        i++;
                    }
                }                
            }

            if (inputList.Count == 0)
                return;

            // Remove random elements till we are at the desired count
            if (order == CompileOrder.Random)
            {
                Random random = new Random(unchecked((int)DateTime.Now.Ticks));

                // Adjust the number of elements in the list
                if (!all)
                {
                    if (maxCount < inputList.Count)
                    {
                        // Randomly prune the list to the desired size

                        int currentCount = inputList.Count;
                        while (currentCount - maxCount > 0)
                        {
                            inputList.RemoveAt(random.Next(currentCount));

                            currentCount--;
                        }
                    }
                    else if (minCount > inputList.Count)
                    {
                        // Randomly duplicate elements in the list to grow it
                        int add = minCount - inputList.Count;
                        int currentCount = inputList.Count;

                        while (add > 0)
                        {
                            inputList.Add(inputList[random.Next(currentCount)]);

                            add--;
                        }
                    }
                }

                // Now randomize the result
                int x = inputList.Count;
                for (int i = 0; i < x; i++)
                {
                    int swapIndex = random.Next(x - i) + i;
                    ListItem swapObject = inputList[i];
                    inputList[i] = inputList[swapIndex];
                    inputList[swapIndex] = swapObject;
                }
            }
            else
            {
                if (order == CompileOrder.Sorted)
                    inputList.Sort();

                if (order == CompileOrder.DailyInterval)
                {
                    TimeSpan ts = currentDate.Subtract(seedDate);

                    int offset = ts.Days % inputList.Count;
 
                    if (offset != 0)
                    {
                        for (int i = 0; i < offset; i++)
                        {
                            inputList.Add(inputList[0]);
                            inputList.RemoveAt(0);
                        }
                    }
                }

                if (!all)
                {
                    if (maxCount < inputList.Count)
                    {
                        inputList.RemoveRange(maxCount, inputList.Count - maxCount);
                    }
                    else if (minCount > inputList.Count)
                    {
                        int currentCount = inputList.Count;
                        int currentIndex = 0;
                        while (inputList.Count < minCount)
                        {
                            if (currentIndex > currentCount)
                                currentIndex = 0;
                            inputList.Add(inputList[currentIndex]);
                            currentIndex++;
                        }
                    }
                }
            }
        }
        
    }
}
