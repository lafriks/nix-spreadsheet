/*
 * Library for writing OLE 2 Compount Document file format.
 * Copyright (C) 2007, Lauris Bukðis-Haberkorns <lauris@nix.lv>
 * 
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * 
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 */

using System;
using System.Collections;

namespace Nix.CompoundFile
{
	/// <summary>
	/// RedBlackTree Enumerator.
	/// </summary>
	internal class RedBlackEnumerator : IEnumerator
	{
        #region Private variables and constructor
        // the treap uses the stack to order the nodes
        private Stack stack;

        // return in ascending order (true) or descending (false)
        private bool ascending;
		
        // root node (needed for reset)
        private Ole2DirectoryEntry root = null;

        // current node
        private Ole2DirectoryEntry current = null;

        /// <summary>
        /// Determine order, walk the tree and push the nodes onto the stack.
        /// </summary>
        /// <param name="tnode">Node to start with.</param>
        /// <param name="ascending">Sort order</param>
        public RedBlackEnumerator(Ole2DirectoryEntry tnode, bool ascending) 
        {
            this.ascending = ascending;
            this.root = tnode;
            this.Reset();
        }
        #endregion

        #region IEnumerator Members

        public void Reset()
        {
            this.stack = new Stack();

            Ole2DirectoryEntry tnode = this.root;

            // use depth-first traversal to push nodes into stack
            // the lowest node will be at the top of the stack
            if (ascending)
            {   // find the lowest node
                while (tnode != RedBlackTree.SentinelNode)
                {
                    this.stack.Push(tnode);
                    tnode = tnode.Left;
                }
            }
            else
            {
                // the highest node will be at top of stack
                while (tnode != RedBlackTree.SentinelNode)
                {
                    this.stack.Push(tnode);
                    tnode = tnode.Right;
                }
            }

            this.current = null;
        }

        public object Current
        {
            get
            {
                if (this.current == null)
                    throw new InvalidOperationException();
                return this.current;
            }
        }

        public bool MoveNext()
        {
            if (stack.Count == 0)
                return false;
			
            // the top of stack will always have the next item
            // get top of stack but don't remove it as the next nodes in sequence
            // may be pushed onto the top
            // the stack will be popped after all the nodes have been returned

            //next node in sequence
            Ole2DirectoryEntry node = (Ole2DirectoryEntry)this.stack.Peek();
			
            if (this.ascending)
            {
                if (node.Right == RedBlackTree.SentinelNode)
                {	
                    // yes, top node is lowest node in subtree - pop node off stack 
                    Ole2DirectoryEntry tn = (Ole2DirectoryEntry)this.stack.Pop();
                    // peek at right node's parent 
                    // get rid of it if it has already been used
                    while(this.stack.Count > 0 && ((Ole2DirectoryEntry)this.stack.Peek()).Right == tn)
                        tn = (Ole2DirectoryEntry)this.stack.Pop();
                }
                else
                {
                    // find the next items in the sequence
                    // traverse to left; find lowest and push onto stack
                    Ole2DirectoryEntry tn = node.Right;
                    while (tn != RedBlackTree.SentinelNode)
                    {
                        this.stack.Push(tn);
                        tn = tn.Left;
                    }
                }
            }
            else
            {
                // descending, same comments as above apply
                if (node.Left == RedBlackTree.SentinelNode)
                {
                    // walk the tree
                    Ole2DirectoryEntry tn = (Ole2DirectoryEntry)this.stack.Pop();
                    while(stack.Count > 0 && ((Ole2DirectoryEntry)this.stack.Peek()).Left == tn)
                        tn = (Ole2DirectoryEntry)this.stack.Pop();
                }
                else
                {
                    // determine next node in sequence
                    // traverse to left subtree and find greatest node - push onto stack
                    Ole2DirectoryEntry tn = node.Left;
                    while (tn != RedBlackTree.SentinelNode)
                    {
                        this.stack.Push(tn);
                        tn = tn.Right;
                    }
                }
            }

            this.current = node;
            return true;
        }

        #endregion
    }
}
