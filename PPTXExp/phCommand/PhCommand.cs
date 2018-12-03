using System;
namespace PPTXExp.phCommand {
    public abstract class PhCommand {
        public virtual Object Exec(params Object[] parameters) {
            Console.WriteLine("something should be exec!");
            throw new Exception("should not abstract command");
        }
    }
}
