/*
 * Library for writing OLE 2 Compount Document file format.
 * Copyright (C) 2007, Lauris Buk�s-Haberkorns <lauris@nix.lv>
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
    internal class SentinelEntry : Ole2DirectoryEntry
    {
        public SentinelEntry() : base(-1)
        {
            this.Left = this;
            this.Right = this;
            this.Parent = null;
            this.Color = NodeColor.Black;
        }
    }
	/// <summary>
	/// RedBlack tree implementation.
	/// </summary>
	internal class RedBlackTree : Object, IEnumerable
	{
        #region Private variables
        /// <summary>
        /// The number of nodes contained in the tree.
        /// </summary>
        private int count = 0;

        /// <summary>
        /// The tree
        /// </summary>
        private Ole2DirectoryEntry rbTree;

        /// <summary>
        /// The node that was last found; used to optimize searches
        /// </summary>
        private Ole2DirectoryEntry lastNodeFound;
        #endregion

        #region Sentinel node
        //  sentinelNode is convenient way of indicating a leaf node.
        private static Ole2DirectoryEntry sentinelNode;          

        public static Ole2DirectoryEntry SentinelNode
        {
            get
            {
                return sentinelNode;
            }
        }

        static RedBlackTree()
        {
            // set up the sentinel node. the sentinel node is the key to a successfull
            // implementation and for understanding the red-black tree properties.
            sentinelNode        = new SentinelEntry();
            /*sentinelNode.Left   = sentinelNode;
            sentinelNode.Right  = sentinelNode;
            sentinelNode.Parent = null;
            sentinelNode.Color  = NodeColor.Black;*/
        }
        #endregion

        #region Constructor
        public RedBlackTree()
        {
            this.rbTree         = SentinelNode;
            this.lastNodeFound  = SentinelNode;
        }
        #endregion

        #region Add node
        /// <summary>
        /// Adds node to the tree.
        /// </summary>
        /// <param name="node">Node to add.</param>
        public void Add(Ole2DirectoryEntry node)
        {
            if(node == null)
                throw new ArgumentNullException("Directory entry must not be null");
			
            // traverse tree - find where node belongs
            int result = 0;
            // grab the rbTree node of the tree
            Ole2DirectoryEntry temp = this.rbTree;

            // find Parent
            while (temp != SentinelNode)
            {
                node.Parent = temp;
                result = node.DID.CompareTo(temp.DID);
                if (result == 0)
                    throw new ArgumentException("A Node with the same DID already exists");
                if (result > 0)
                    temp = temp.Right;
                else
                    temp = temp.Left;
            }

            // setup node
            node.Left  = SentinelNode;
            node.Right = SentinelNode;

            // insert node into tree starting at parent's location
            if (node.Parent != null)	
            {
                result = node.DID.CompareTo(node.Parent.DID);
                if (result > 0)
                    node.Parent.Right = node;
                else
                    node.Parent.Left = node;
            }
            else
            {
                // first node added
                this.rbTree = node;
            }

            // restore red-black properities
            this.RestoreAfterInsert(node);

            this.lastNodeFound = node;

            this.count++;
        }
        #endregion

        #region Rotation and color restoration functions
        /// <summary>
        /// Additions to red-black trees usually destroy the red-black 
        /// properties. Examine the tree and restore. Rotations are normally 
        /// required to restore it.
        /// </summary>
        /// <param name="x">Node to start from</param>
        private void RestoreAfterInsert(Ole2DirectoryEntry x)
        {   
            Ole2DirectoryEntry y;

            // maintain red-black tree properties after adding x
            while (x != this.rbTree && x.Parent.Color == NodeColor.Red)
            {
                // Parent node is .Colored red; 
                if(x.Parent == x.Parent.Parent.Left)	// determine traversal path			
                {										// is it on the Left or Right subtree?
                    y = x.Parent.Parent.Right;			// get uncle
                    if (y.Color == NodeColor.Red) // y != null && ???
                    {
                        // uncle is red; change x's Parent and uncle to black
                        x.Parent.Color			= NodeColor.Black;
                        y.Color 				= NodeColor.Black;
                        // grandparent must be red. Why? Every red node that is not 
                        // a leaf has only black children 
                        x.Parent.Parent.Color	= NodeColor.Red;
                        // continue loop with grandparent
                        x = x.Parent.Parent;
                    }
                    else
                    {
                        // uncle is black; determine if x is greater than Parent
                        if (x == x.Parent.Right) 
                        {	// yes, x is greater than Parent; rotate Left
                            // make x a Left child
                            x = x.Parent;
                            this.RotateLeft(x);
                        }
                        // TODO: Shouln't there be else?
                        // no, x is less than Parent
                        // make Parent black
                        x.Parent.Color			= NodeColor.Black;
                        // make grandparent black
                        x.Parent.Parent.Color	= NodeColor.Red;
                        // rotate right
                        this.RotateRight(x.Parent.Parent);
                    }
                }
                else
                {
                    // x's Parent is on the Right subtree
                    // this code is the same as above with "Left" and "Right" swapped
                    y = x.Parent.Parent.Left;
                    if (y.Color == NodeColor.Red) // y != null && ???
                    {
                        x.Parent.Color			= NodeColor.Black;
                        y.Color					= NodeColor.Black;
                        x.Parent.Parent.Color	= NodeColor.Red;
                        // continue loop with grandparent
                        x = x.Parent.Parent;
                    }
                    else
                    {
                        if(x == x.Parent.Left)
                        {
                            x = x.Parent;
                            this.RotateRight(x);
                        }
                        // TODO: Shouln't there be else?
                        x.Parent.Color			= NodeColor.Black;
                        x.Parent.Parent.Color	= NodeColor.Red;
                        // rotate left
                        this.RotateLeft(x.Parent.Parent);
                    }
                }																													
            }
            // rbTree should always be black
            this.rbTree.Color = NodeColor.Black;
        }
		
        /// <summary>
        /// Rebalance the tree by rotating the nodes to the left.
        /// </summary>
        /// <param name="x">Node to start with.</param>
        private void RotateLeft(Ole2DirectoryEntry x)
        {
            // pushing node x down and to the Left to balance the tree. x's Right child (y)
            // replaces x (since y > x), and y's Left child becomes x's Right child 
            // (since it's < y but > x).

            // get x's Right node, this becomes y
            Ole2DirectoryEntry y = x.Right;

            // set x's Right link
            // y's Left child's becomes x's Right child
            x.Right = y.Left;

            // modify parents
            if (y.Left != SentinelNode)
            {
                // sets y's Left Parent to x
                y.Left.Parent = x;
            }

            if (y != SentinelNode)
            {
                // set y's Parent to x's Parent
                y.Parent = x.Parent;
            }

            if (x.Parent != null)		
            {
                // determine which side of it's Parent x was on
                if (x == x.Parent.Left)			
                    x.Parent.Left = y;	// set Left Parent to y
                else
                    x.Parent.Right = y;	// set Right Parent to y
            }
            else 
            {
                // at rbTree, set it to y
                this.rbTree = y;
            }

            // link x and y 
            // put x on y's Left 
            y.Left = x;
            if (x != SentinelNode)
            {
                // set y as x's Parent
                x.Parent = y;
            }
        }

        /// <summary>
        /// Rebalance the tree by rotating the nodes to the right.
        /// </summary>
        /// <param name="x">Node to start with.</param>
        private void RotateRight(Ole2DirectoryEntry x)
        {
            // pushing node x down and to the Right to balance the tree. x's Left child (y)
            // replaces x (since x < y), and y's Right child becomes x's Left child 
            // (since it's < x but > y).

            // get x's Left node, this becomes y
            Ole2DirectoryEntry y = x.Left;

            // set x's Right link
            // y's Right child becomes x's Left child
            x.Left = y.Right;

            // modify parents
            if (y.Right != SentinelNode)
            {
                // sets y's Right Parent to x
                y.Right.Parent = x;
            }

            if (y != SentinelNode)
            {
                // set y's Parent to x's Parent
                y.Parent = x.Parent;
            }

            // null=rbTree, could also have used rbTree
            if (x.Parent != null)
            {
                // determine which side of it's Parent x was on
                if (x == x.Parent.Right)			
                    x.Parent.Right = y;	// set Right Parent to y
                else
                    x.Parent.Left = y;	// set Left Parent to y
            } 
            else
            {
                // at rbTree, set it to y
                this.rbTree = y;
            }

            // link x and y 
            // put x on y's Right
            y.Right = x;
            if (x != SentinelNode)
            {
                // set y as x's Parent
                x.Parent = y;
            }
        }
        #endregion

        #region Get
        /// <summary>
        /// Gets the directory entry with the specified DID in the tree.
        /// </summary>
        public Ole2DirectoryEntry this[int DID]
        {
            get
            {
                int result = 0;

                // begin at root
                Ole2DirectoryEntry treeNode = this.rbTree;

                // traverse tree until node is found
                while (treeNode != SentinelNode)
                {
                    result = DID.CompareTo(treeNode.DID);
                    if (result == 0)
                    {
                        this.lastNodeFound = treeNode;
                        return treeNode;
                    }
                    if (result < 0)
                        treeNode = treeNode.Left;
                    else
                        treeNode = treeNode.Right;
                }

                throw(new ArgumentException("Directory entry with such DID was not found"));
            }
        }

        /// <summary>
        /// Returns the minimum DID in the tree
        /// </summary>
        /// <returns>Minimum DID</returns>
        public int GetMinDID()
        {
            Ole2DirectoryEntry treeNode = this.rbTree;
			
            if (this.IsEmpty())
                throw(new ArgumentException("RedBlack tree is empty"));
			
            // traverse to the extreme left to find the smallest key
            while (treeNode.Left != SentinelNode)
                treeNode = treeNode.Left;
			
            this.lastNodeFound = treeNode;
			
            return treeNode.DID;
			
        }
        /// <summary>
        /// Returns the maximum DID in the tree
        /// </summary>
        /// <returns>Maximum DID</returns>
        public int GetMaxDID()
        {
            Ole2DirectoryEntry treeNode = this.rbTree;
			
            if (this.IsEmpty())
                throw(new ArgumentException("RedBlack tree is empty"));

            // traverse to the extreme right to find the largest key
            while (treeNode.Right != SentinelNode)
                treeNode = treeNode.Right;

            this.lastNodeFound = treeNode;

            return treeNode.DID;
			
        }

        /// <summary>
        /// Returns the directory entry having the minimum DID.
        /// </summary>
        /// <returns>Directory entry.</returns>
        public Ole2DirectoryEntry GetMinNode()
        {
            return this[this.GetMinDID()];
        }

        /// <summary>
        /// Returns the directory entry having the maximum DID.
        /// </summary>
        /// <returns>Directory entry.</returns>
        public Ole2DirectoryEntry GetMaxNode()
        {
            return this[this.GetMaxDID()];
        }
        #endregion

        #region Remove node
        /// <summary>
        /// Removes the node from the tree.
        /// </summary>
        /// <param name="key">Directory DID</param>
        public void Remove(int DID)
        {
            // find node
            Ole2DirectoryEntry node;

            // see if node to be deleted was the last one found
            if (this.lastNodeFound.DID == DID)
                node = this.lastNodeFound;
            else
            {
                // not found, must search
                node = this.rbTree;
                int result;
                while (node != SentinelNode)
                {
                    result = DID.CompareTo(node.DID);
                    if (result == 0)
                        break;
                    if (result < 0)
                        node = node.Left;
                    else
                        node = node.Right;
                }

                // key not found
                if (node == SentinelNode)
                    return;
            }

            Delete(node);
            count--;
        }

        /// <summary>
        /// Delete a node from the tree and restore red black properties
        /// </summary>
        /// <param name="z">Node to delete.</param>
        private void Delete(Ole2DirectoryEntry z)
        {
            // A node to be deleted will be: 
            //		1. a leaf with no children
            //		2. have one child
            //		3. have two children
            // If the deleted node is red, the red black properties still hold.
            // If the deleted node is black, the tree needs rebalancing

            // work node to contain the replacement node
            Ole2DirectoryEntry x;
            // work node 
            Ole2DirectoryEntry y;

            // find the replacement node (the successor to x) - the node one with 
            // at *most* one child. 
            if(z.Left == SentinelNode || z.Right == SentinelNode)
            {
                // node has sentinel as a child
                y = z;
            }
            else 
            {
                // z has two children, find replacement node which will 
                // be the leftmost node greater than z
                y = z.Right;				        // traverse right subtree	
                while (y.Left != SentinelNode)		// to find next node in sequence
                    y = y.Left;
            }

            // at this point, y contains the replacement node. it's content will be copied 
            // to the valules in the node to be deleted

            // x (y's only child) is the node that will be linked to y's old parent. 
            if (y.Left != SentinelNode)
                x = y.Left;					
            else
                x = y.Right;					

            // replace x's parent with y's parent and
            // link x to proper subtree in parent
            // this removes y from the chain
            x.Parent = y.Parent;
            if (y.Parent != null)
            {
                if (y == y.Parent.Left)
                    y.Parent.Left = x;
                else
                    y.Parent.Right = x;
            }
            else
            {
                // make x the root node
                this.rbTree = x;
            }

            // copy the values from y (the replacement node) to the node being deleted.
            // note: this effectively deletes the node. ???
            //if (y != z) 
            //{
            //    z.DID = y.DID;
            //}

            if (y.Color == NodeColor.Black)
                this.RestoreAfterDelete(x);

            this.lastNodeFound = SentinelNode;
        }

        /// <summary>
        /// Deletions from red-black trees may destroy the red-black 
        /// properties. Examine the tree and restore. Rotations are normally 
        /// required to restore it
        /// </summary>
        /// <param name="x">Node to start with.</param>
        private void RestoreAfterDelete (Ole2DirectoryEntry x)
        {
            // maintain Red-Black tree balance after deleting node 			
            Ole2DirectoryEntry y;

            while (x != this.rbTree && x.Color == NodeColor.Black) 
            {
                // determine sub tree from parent
                if (x == x.Parent.Left)
                {
                    // y is x's sibling 
                    y = x.Parent.Right;
                    if (y.Color == NodeColor.Red) 
                    {	// x is black, y is red - make both black and rotate
                        y.Color			= NodeColor.Black;
                        x.Parent.Color	= NodeColor.Red;
                        this.RotateLeft(x.Parent);
                        y = x.Parent.Right;
                    }
                    if (y.Left.Color == NodeColor.Black && 
                        y.Right.Color == NodeColor.Black) 
                    {
                        // children are both black
                        // change parent to red
                        y.Color = NodeColor.Red;
                        // move up the tree
                        x = x.Parent;
                    } 
                    else 
                    {
                        if (y.Right.Color == NodeColor.Black) 
                        {
                            y.Left.Color    = NodeColor.Black;
                            y.Color			= NodeColor.Red;
                            this.RotateRight(y);
                            y = x.Parent.Right;
                        }
                        y.Color = x.Parent.Color;
                        x.Parent.Color	= NodeColor.Black;
                        y.Right.Color	= NodeColor.Black;
                        this.RotateLeft(x.Parent);
                        x = this.rbTree;
                    }
                } 
                else 
                {	// right subtree - same as code above with right and left swapped
                    y = x.Parent.Left;
                    if (y.Color == NodeColor.Red) 
                    {
                        y.Color			= NodeColor.Black;
                        x.Parent.Color	= NodeColor.Red;
                        this.RotateRight(x.Parent);
                        y = x.Parent.Left;
                    }
                    if (y.Right.Color == NodeColor.Black && 
                        y.Left.Color == NodeColor.Black) 
                    {
                        y.Color = NodeColor.Red;
                        x = x.Parent;
                    }
                    else
                    {
                        if (y.Left.Color == NodeColor.Black) 
                        {
                            y.Right.Color	= NodeColor.Black;
                            y.Color			= NodeColor.Red;
                            this.RotateLeft(y);
                            y = x.Parent.Left;
                        }
                        y.Color = x.Parent.Color;
                        x.Parent.Color	= NodeColor.Black;
                        y.Left.Color	= NodeColor.Black;
                        this.RotateRight(x.Parent);
                        x = this.rbTree;
                    }
                }
            }
            x.Color = NodeColor.Black;
        }

        /// <summary>
        /// Removes the node with the minimum DID
        /// </summary>
        public void RemoveMin()
        {
            this.Remove(this.GetMinDID());
        }
        /// <summary>
        /// Removes the node with the maximum DID
        /// </summary>
        public void RemoveMax()
        {
            this.Remove(this.GetMaxDID());
        }
        #endregion

        #region Misc methods
        /// <summary>
        /// Is the tree empty?
        /// </summary>
        /// <returns>true if tree is emty.</returns>
        public bool IsEmpty()
        {
            return (this.rbTree == SentinelNode);
        }

        /// <summary>
        /// Empties or clears the tree
        /// </summary>
        public void Clear()
        {
            this.rbTree = SentinelNode;
            this.count = 0;
        }

        /// <summary>
        /// Returns the size (number of nodes) in the tree
        /// </summary>
        /// <returns>Number of nodes.</returns>
        public int Count()
        {
            // number of keys
            return this.count;
        }

        /// <summary>
        /// Returns root node of the tree
        /// </summary>
        public Ole2DirectoryEntry Root
        {
            get
            {
                return this.rbTree;
            }
        }
        #endregion

        #region IEnumerable Members
        /// <summary>
        /// Keys in ascending order.
        /// </summary>
        /// <returns>RedBlackTree enumerator</returns>
        public RedBlackEnumerator Keys()
        {
            return this.Keys(true);
        }
        /// <summary>
        /// Keys in specified order.
        /// </summary>
        /// <param name="ascending">Order.</param>
        /// <returns>RedBlackTree enumerator</returns>
        public RedBlackEnumerator Keys(bool ascending)
        {
            return new RedBlackEnumerator(this.rbTree,  ascending);
        }

        public IEnumerator GetEnumerator()
        {
            return this.Keys(true);
        }
        #endregion
    }
}
