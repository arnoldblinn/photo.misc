using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Text;

namespace Msn.PhotoMix.SlideShow
{
    public class FriendlyName
    {
        static public bool CheckFriendlyNameAvailable(string friendlyName)
        {
            string sql = "select FriendlyName from FriendlyNames where FriendlyName = @FriendlyName and FriendlyNameHash = @FriendlyNameHash";
            using (PhotoMixQuery query = new PhotoMixQuery(sql, System.Data.CommandType.Text))
            {
                query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = friendlyName;
                query.Parameters.Add("@FriendlyNameHash", SqlDbType.Int).Value = friendlyName.GetHashCode();

                if (query.Reader.Read())
                    return false;
            }

            return true;
        }

        //
        // DeleteFriendlyName
        //
        // Removes the friendly name for the passed in slideshow
        //
        static public void DeleteFriendlyName(string friendlyName, Guid slideShowGuid, int puidHash)
        {
            // First delete the friendly name from the lookup table            
            if (!String.IsNullOrEmpty(friendlyName))
            {
                string sql = "delete from FriendlyNames where FriendlyName = @FriendlyName and SlideShowGuid = @SlideShowGuid and PuidHash = @PuidHash and FriendlyNameHash = @FriendlyNameHash";
                using (PhotoMixQuery query = new PhotoMixQuery(sql, System.Data.CommandType.Text))
                {
                    query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = friendlyName;
                    query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = slideShowGuid;
                    query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puidHash;
                    query.Parameters.Add("@FriendlyNameHash", SqlDbType.Int).Value = friendlyName.GetHashCode();

                    query.Execute();
                }
            }
        }

        //
        // UpdateFriendlyName
        //
        // Updates the frienldy name from the old friendly name to new friendly name for the passed in slideshow.  
        //
        static public void UpdateFriendlyName(string newFriendlyName, string oldFriendlyName, Guid slideShowGuid, int puidHash)
        {
            UpdateFriendlyName(newFriendlyName, oldFriendlyName, slideShowGuid, puidHash, Guid.Empty);
        }

        //
        // UpdateFriendlyName
        //
        // Updates the friendly name from the old friendly name to the new friendly name for the passed in slideshow.  If the old slideshow is
        // specified (optional) than the new friendly name is currently assigned to this slideshow and will be reasigned to the new slide show.  It
        // is assumed that this old slideshow is owned by the same PUID as the new slideshow
        //
        static public void UpdateFriendlyName(string newFriendlyName, string oldFriendlyName, Guid slideShowGuid, int puidHash, Guid oldSlideShowGuid)
        {
            string sql;

            // We need to do this as distinct database transactions for future web store 
            // integration.  This is because the friendly name partition is different than the slide show partition.  As such
            // updates to the friendly name table and the slideshow table would likely talk to different databases.
            //
            // To handle this we assume that the friendly name table is the definitive source of ownership of a friendly name.
            // The slideshow table has the friendly name, but this is verfied before it is assumed to be valid.
            //
            // By doing so the right way we are never left in an inconsistent state.  The Slideshow
            // itself can deal with thinking it has a friendly name it doesn't own, but not the other way around
            // So we order these transactions such if there is a failure of a subsequent write in this sequence things
            // are not left in a bad state or with dangling data that won't be cleaned up

            // First delete the old friendly name from the friendly name table            
            if (!String.IsNullOrEmpty(oldFriendlyName))
            {
                sql = "delete from FriendlyNames " +
                        "where FriendlyName = @FriendlyName and SlideShowGuid = @SlideShowGuid and PuidHash = @PuidHash and FriendlyNameHash = @FriendlyNameHash";
                using (PhotoMixQuery query = new PhotoMixQuery(sql, System.Data.CommandType.Text))
                {
                    query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = oldFriendlyName;
                    query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = slideShowGuid;
                    query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puidHash; 
                    query.Parameters.Add("@FriendlyNameHash", SqlDbType.Int).Value = oldFriendlyName.GetHashCode();

                    query.Execute();
                }
            }

            // Now update the slide show friendly name to the new value
            sql = "" +
                "update SlideShows " +
                "set FriendlyName = @FriendlyName " +
                "where SlideShowGuid = @SlideShowGuid and PuidHash = @PuidHash";
            using (PhotoMixQuery query = new PhotoMixQuery(sql, System.Data.CommandType.Text))
            {
                if (String.IsNullOrEmpty(newFriendlyName))
                    query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = DBNull.Value;
                else
                    query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = newFriendlyName;
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puidHash;
                query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = slideShowGuid;
                query.Execute();
            }

            // If we are taking the name from another 
            if (oldSlideShowGuid != Guid.Empty)
            {
                // Update the friendly name table for this friendly name
                sql = "update FriendlyNames " +
                        "set SlideShowGuid = @SlideShowGuid, PuidHash = @PuidHash " +
                        "where FriendlyName = @FriendlyName and FriendlyNameHash = @FriendlyNameHash";
                using (PhotoMixQuery query = new PhotoMixQuery(sql, System.Data.CommandType.Text))
                {                    
                    query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puidHash;
                    query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = slideShowGuid;
                    query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = newFriendlyName;
                    query.Parameters.Add("@FriendlyNameHash", SqlDbType.Int).Value = newFriendlyName.GetHashCode();

                    query.Execute();
                }

                // Clear the friendly name from the old slideshow
                sql = "" +
                    "update SlideShows " +
                    "set FriendlyName = @FriendlyName " +
                    "where SlideShowGuid = @SlideShowGuid and PuidHash = @PuidHash";
                using (PhotoMixQuery query = new PhotoMixQuery(sql, System.Data.CommandType.Text))
                {                    
                    query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = DBNull.Value;                 
                    query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puidHash;
                    query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = oldSlideShowGuid;
                    query.Execute();
                }
            }
            // We are simply assigning a new friendly name
            else
            {
                // Insert the new friendly name into the lookup table
                if (!String.IsNullOrEmpty(newFriendlyName))
                {
                    sql = "" +
                        "if not exists (select FriendlyName from FriendlyNames where FriendlyName = @FriendlyName and FriendlyNameHash = @FriendlyNameHash)" +
                        "insert into FriendlyNames (" +
                        "   FriendlyName, SlideShowGuid, PuidHash, FriendlyNameHash" +
                        ") values (" +
                        "   @FriendlyName, @SlideShowGuid, @PuidHash, @FriendlyNameHash" +
                        ")";
                    using (PhotoMixQuery query = new PhotoMixQuery(sql, System.Data.CommandType.Text))
                    {
                        query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = newFriendlyName;
                        query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puidHash;
                        query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = slideShowGuid;
                        query.Parameters.Add("@FriendlyNameHash", SqlDbType.Int).Value = newFriendlyName.GetHashCode();

                        query.Execute();
                    }
                }
            }
        }
        
        static public Guid LookupFriendlyName(string friendlyName, out int puidHash)
        {
            string sql = "select SlideShowGuid, PuidHash from FriendlyNames where FriendlyName = @FriendlyName and FriendlyNameHash = @FriendlyNameHash";

            using (PhotoMixQuery query = new PhotoMixQuery(sql, System.Data.CommandType.Text))
            {
                query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = friendlyName;         
                query.Parameters.Add("@FriendlyNameHash", SqlDbType.Int).Value = friendlyName.GetHashCode();

                if (query.Reader.Read())
                {
                    puidHash = query.Reader.GetInt32(1);
                    return query.Reader.GetGuid(0);
                }
            }

            puidHash = 0;
            return Guid.Empty;
        }
    }
}
