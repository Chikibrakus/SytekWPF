﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace СитекWPF.DataFiles
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class FIAS_GAREntities : DbContext
    {
        public FIAS_GAREntities()
            : base("name=FIAS_GAREntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<ADDR_OBJ> ADDR_OBJ { get; set; }
        public virtual DbSet<AS_OBJECT_LEVELS> AS_OBJECT_LEVELS { get; set; }
        public virtual DbSet<sysdiagrams> sysdiagrams { get; set; }
    }
}
