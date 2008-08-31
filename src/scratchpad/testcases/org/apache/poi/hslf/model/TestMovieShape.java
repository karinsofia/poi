/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.apache.poi.hslf.model;

import junit.framework.TestCase;

import java.io.*;
import java.awt.*;
import java.awt.geom.Rectangle2D;

import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.hslf.HSLFSlideShow;

/**
 * Test <code>MovieShape</code> object.
 * 
 * @author Yegor Kozlov
 */
public class TestMovieShape extends TestCase {

    protected String cwd = System.getProperty("HSLF.testdata.path");

    public void testCreate() throws Exception {
        SlideShow ppt = new SlideShow();

        Slide slide = ppt.createSlide();

        String path = cwd + "/test-movie.mpg";
        int movieIdx = ppt.addMovie(path, MovieShape.MOVIE_MPEG);
        int thumbnailIdx = ppt.addPicture(new File(cwd, "tomcat.png"), Picture.PNG);

        MovieShape shape = new MovieShape(movieIdx, thumbnailIdx);
        shape.setAnchor(new Rectangle2D.Float(300,225,120,90));
        slide.addShape(shape);

        assertEquals(path, shape.getPath());
        assertTrue(shape.isAutoPlay());
        shape.setAutoPlay(false);
        assertFalse(shape.isAutoPlay());

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ppt.write(out);

        ppt = new SlideShow(new ByteArrayInputStream(out.toByteArray()));
        slide = ppt.getSlides()[0];
        shape = (MovieShape)slide.getShapes()[0];
        assertEquals(path, shape.getPath());
        assertFalse(shape.isAutoPlay());

    }

}