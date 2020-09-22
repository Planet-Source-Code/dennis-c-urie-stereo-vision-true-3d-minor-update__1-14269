Stereo Vision Custom model definition file format.

For simplicity, the model definition files are plain text files providing
only the necessary information needed for the application.  The program
does NOT attempt to verify the file before attempting to load it, if there
any errors in the file it may cause unexpected side effects or a file read
error.

File Format is as follows:

Line 1, Texture file name in quotes    
Line 2, Total number of faces (330 max) in the model
Line 3, Number of Verts in face (Must be set to 4 per face at this time)
Lines 5 - 8, Vertex information as follows:

  X, Y, Z, NormalX, NormalY, NormalZ, tu, tv, color

  X, Y, Z                    - Position of vertex
  NormalX, NormalY, NormalZ  - Vertex Normal used for gouraud light calculations.
  tu, tv                     - vertex Texture coordinates
  color                      - Long integer representation of vertex color         

Lines 9 and up, Repeat lines 3 - 8 for each additional face.

Comments can be placed at the end of the file.

Notes: Currently the program only has support for 4 sided polygons, A triangle 
       is created by assigning the first or last two verts to the same 
       coordinates.

       For the face to be drawn to the screen the verts must be oriented in a 
       clockwise fashion when the face is facing towards the viewer.

       In order for backface culling and flat shading to function properly all
       four verts must be on the same plane.

       Objects are rotated on the origin 0,0,0

       Vertex normals are only used for light calculations, and do NOT need to
       based from the origin.